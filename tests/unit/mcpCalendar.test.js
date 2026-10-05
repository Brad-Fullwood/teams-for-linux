'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

process.env.TZ = 'Europe/London';
const { renderEventLine, renderWhen, parseGraphTime, parseStart } = require('../../app/mcp/tools/calendar');

describe('calendar', () => {
	it('reads Graph UTC times that have no zone suffix', () => {
		assert.strictEqual(parseGraphTime({ dateTime: '2026-10-05T08:30:00.0000000', timeZone: 'UTC' }).toISOString(), '2026-10-05T08:30:00.000Z');
		assert.strictEqual(parseGraphTime(undefined), null);
	});

	it('treats a bare date as local midnight', () => {
		assert.strictEqual(parseStart('2026-10-05').toISOString(), '2026-10-04T23:00:00.000Z');
		assert.strictEqual(parseStart('nonsense'), null);
	});

	it('renders timed events in local time', () => {
		const when = renderWhen({ start: { dateTime: '2026-10-05T08:30:00.0000000', timeZone: 'UTC' }, end: { dateTime: '2026-10-05T09:00:00.0000000', timeZone: 'UTC' } });
		assert.strictEqual(when, 'Mon 2026-10-05 09:30-10:00');
	});

	it('renders all-day events as dates with an exclusive end', () => {
		assert.strictEqual(renderWhen({ isAllDay: true, start: { dateTime: '2026-10-05T00:00:00.0000000' }, end: { dateTime: '2026-10-06T00:00:00.0000000' } }), '2026-10-05, all day');
		assert.strictEqual(renderWhen({ isAllDay: true, start: { dateTime: '2026-10-05T00:00:00.0000000' }, end: { dateTime: '2026-10-08T00:00:00.0000000' } }), '2026-10-05 to 2026-10-07, all day');
	});

	it('renders an event line with flags, organiser, link and id', () => {
		const line = renderEventLine({
			id: 'EV1', subject: 'Design review', showAs: 'tentative', type: 'occurrence',
			start: { dateTime: '2026-10-08T13:00:00.0000000', timeZone: 'UTC' }, end: { dateTime: '2026-10-08T14:00:00.0000000', timeZone: 'UTC' },
			organizer: { emailAddress: { name: 'Sam', address: 'sam@x.com' } }, responseStatus: { response: 'accepted' },
			onlineMeeting: { joinUrl: 'https://teams/join' }, bodyPreview: 'Agenda:\n  scope\n________________________________ Microsoft Teams meeting Join',
		});
		assert.strictEqual(line, [
			'- [Thu 2026-10-08 14:00-15:00] Design review [TENTATIVE,RECURRING]',
			'  organiser: Sam <sam@x.com>, my response: accepted',
			'  teams: https://teams/join',
			'  preview: Agenda: scope',
			'  id: EV1',
		].join('\n'));
	});

	it('drops a preview that is only the Teams join block', () => {
		const line = renderEventLine({ id: 'EV2', subject: 'x', start: { dateTime: '2026-10-08T13:00:00', timeZone: 'UTC' }, bodyPreview: '______________________________ Microsoft Teams meeting' });
		assert.ok(!line.includes('preview:'));
	});
});
