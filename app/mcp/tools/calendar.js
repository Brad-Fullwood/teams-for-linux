const { z } = require('zod');
const { htmlToText, truncate, toolResult, personName, personAddress } = require('./format');

const EVENT_SELECT = 'id,subject,start,end,isAllDay,isCancelled,showAs,organizer,location,onlineMeeting,isOnlineMeeting,responseStatus,categories,bodyPreview,webLink,seriesMasterId,type';

/** Graph returns UTC times without a zone suffix when no Prefer: outlook.timezone header is sent. */
function parseGraphTime(value) {
  if (!value?.dateTime) return null;
  const base = value.dateTime.slice(0, 19);
  const zoned = !value.timeZone || value.timeZone === 'UTC' ? `${base}Z` : base;
  const d = new Date(zoned);
  return Number.isNaN(d.getTime()) ? null : d;
}

/** Local midnight for a YYYY-MM-DD string, or the instant for a full ISO date-time. */
function parseStart(value) {
  if (!value) {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }
  const dateOnly = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  const d = dateOnly ? new Date(Number(dateOnly[1]), Number(dateOnly[2]) - 1, Number(dateOnly[3])) : new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

const pad = (n) => String(n).padStart(2, '0');
const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const localDay = (d) => `${DAY_NAMES[d.getDay()]} ${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const localTime = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;

/** "Mon 2026-10-05 09:30-10:00" in local time; all-day events show their dates only. */
function renderWhen(event) {
  if (event.isAllDay) {
    // All-day dates are calendar dates, not instants: read them as written.
    const first = event.start?.dateTime?.slice(0, 10);
    const endExclusive = event.end?.dateTime?.slice(0, 10);
    const last = endExclusive ? addDays(parseStart(endExclusive), -1) : null;
    const lastText = last ? `${last.getFullYear()}-${pad(last.getMonth() + 1)}-${pad(last.getDate())}` : first;
    return lastText && lastText !== first ? `${first} to ${lastText}, all day` : `${first}, all day`;
  }
  const start = parseGraphTime(event.start);
  const end = parseGraphTime(event.end);
  if (!start) return '(no start time)';
  if (!end) return `${localDay(start)} ${localTime(start)}`;
  const sameDay = localDay(start) === localDay(end);
  return `${localDay(start)} ${localTime(start)}-${sameDay ? '' : `${localDay(end)} `}${localTime(end)}`;
}

// Teams appends its join block below a line of underscores. Keep only what the organiser wrote.
const JOIN_BLOCK = /_{20,}[\s\S]*$/;
const meetingText = (text) => (text ?? '').replace(JOIN_BLOCK, '').trim();

function renderEventLine(e) {
  const flags = [
    e.isCancelled ? 'CANCELLED' : '',
    e.showAs && e.showAs !== 'busy' ? e.showAs.toUpperCase() : '',
    e.type === 'occurrence' || e.type === 'exception' ? 'RECURRING' : '',
  ].filter(Boolean).join(',');
  const response = e.responseStatus?.response;
  const lines = [
    `- [${renderWhen(e)}] ${e.subject || '(no subject)'}${flags ? ` [${flags}]` : ''}`,
    `  organiser: ${personName(e.organizer)}${personAddress(e.organizer)}${response && response !== 'none' ? `, my response: ${response}` : ''}`,
  ];
  const where = e.location?.displayName;
  if (where) lines.push(`  where: ${where}`);
  if (e.onlineMeeting?.joinUrl) lines.push(`  teams: ${e.onlineMeeting.joinUrl}`);
  if (e.categories?.length) lines.push(`  categories: ${e.categories.join(', ')}`);
  const preview = truncate(meetingText(e.bodyPreview).replaceAll(/\s+/g, ' ').trim(), 160);
  if (preview) lines.push(`  preview: ${preview}`);
  lines.push(`  id: ${e.id}`);
  return lines.join('\n');
}

/**
 * Register read-only calendar tools.
 * @param {import('@modelcontextprotocol/sdk/server/mcp.js').McpServer} server
 * @param {object} graph - GraphApiClient
 */
function registerCalendarTools(server, graph) {
  server.registerTool('calendar_list_events', {
    title: 'List calendar events',
    description: 'List events in the signed-in user\'s Outlook calendar for a date range, in start order, recurring meetings expanded. Times are shown in this machine\'s local time zone. Includes all-day events (often used for bookings and leave), cancelled and tentative markers, organiser, Teams join link and ids for calendar_get_event.',
    inputSchema: {
      start: z.string().optional().describe('YYYY-MM-DD (local midnight) or an ISO date-time. Default: today'),
      days: z.number().int().min(1).max(62).default(1).describe('Number of days from start'),
      limit: z.number().int().min(1).max(200).default(100),
      includeCancelled: z.boolean().default(false),
    },
  }, async ({ start, days, limit, includeCancelled }) => {
    const from = parseStart(start);
    if (!from) return { isError: true, content: [{ type: 'text', text: `Invalid start: ${start}` }] };
    const to = addDays(from, days);
    const result = await graph.getCalendarView(from.toISOString(), to.toISOString(), {
      top: limit, select: EVENT_SELECT, orderby: 'start/dateTime',
    });
    return toolResult(result, (data) => {
      const items = (data?.value ?? []).filter((e) => includeCancelled || !e.isCancelled);
      const range = `${localDay(from)} to ${localDay(addDays(to, -1))}`;
      if (items.length === 0) return `No events ${range}.`;
      const footer = data?.['@odata.nextLink'] ? '\n(more events in this range; raise limit or narrow the range)' : '';
      return `${items.length} event(s) ${range}:\n${items.map(renderEventLine).join('\n')}${footer}`;
    });
  });

  server.registerTool('calendar_get_event', {
    title: 'Get a calendar event',
    description: 'Read one calendar event in full: time, organiser, attendees with their responses, location, Teams link and the body (agenda) as plain text. Use the id from calendar_list_events.',
    inputSchema: {
      id: z.string(),
      maxBodyChars: z.number().int().min(200).max(50_000).default(8000),
    },
  }, async ({ id, maxBodyChars }) => {
    const result = await graph.makeRequest(`/me/events/${encodeURIComponent(id)}?$select=${EVENT_SELECT},attendees,body`);
    return toolResult(result, (e) => {
      const attendees = (e.attendees ?? []).map((a) => {
        const response = a.status?.response;
        return `${personName(a)}${personAddress(a)} (${a.type ?? 'attendee'}${response && response !== 'none' ? `, ${response}` : ''})`;
      });
      const body = e.body?.contentType === 'html' ? htmlToText(e.body.content) : (e.body?.content ?? '');
      return [
        `Subject: ${e.subject || '(no subject)'}`,
        `When: ${renderWhen(e)}${e.isCancelled ? ' (CANCELLED)' : ''}`,
        `Organiser: ${personName(e.organizer)}${personAddress(e.organizer)}`,
        `Show as: ${e.showAs ?? 'unknown'}${e.responseStatus?.response && e.responseStatus.response !== 'none' ? `, my response: ${e.responseStatus.response}` : ''}`,
        e.location?.displayName ? `Where: ${e.location.displayName}` : null,
        e.onlineMeeting?.joinUrl ? `Teams: ${e.onlineMeeting.joinUrl}` : null,
        e.categories?.length ? `Categories: ${e.categories.join(', ')}` : null,
        attendees.length ? `Attendees: ${attendees.join('; ')}` : null,
        `Link: ${e.webLink}`,
        '',
        truncate(meetingText(body) || '(no agenda)', maxBodyChars),
      ].filter((line) => line !== null).join('\n');
    });
  });
}

module.exports = { registerCalendarTools, renderEventLine, renderWhen, parseGraphTime, parseStart };
