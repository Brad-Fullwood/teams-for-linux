'use strict';

// Encoding sets below are the ones Teams used in a 2026-10-27 meeting (read
// from the Teams page console) and in 1:1 calls.

const { describe, it, beforeEach, afterEach, mock } = require('node:test');
const assert = require('node:assert');

const modulePath = require.resolve('../../app/browser/tools/cameraSendResolution');

const LIFTED = { maxWidth: 1920, maxHeight: 1080 };
const THUMB = { rid: '1', active: true, maxBitrate: 133200, maxFramerate: 15, scalabilityMode: 'L1T1', scaleResolutionDownBy: 6, scaleResolutionDownTo: { maxWidth: 320, maxHeight: 180 } };
const MAIN_720 = { rid: '2', active: true, maxBitrate: 1536000, maxFramerate: 30, scalabilityMode: 'L1T1', scaleResolutionDownBy: 1.5, scaleResolutionDownTo: { maxWidth: 1280, maxHeight: 720 } };
const MAIN_540 = { ...MAIN_720, maxBitrate: 855600, scaleResolutionDownBy: 2, scaleResolutionDownTo: { maxWidth: 960, maxHeight: 540 } };
const MAIN_180 = { ...MAIN_720, maxBitrate: 126000, maxFramerate: 15, scaleResolutionDownBy: 6, scaleResolutionDownTo: { maxWidth: 320, maxHeight: 180 } };
const UNUSED = { rid: '3', active: false, scaleResolutionDownBy: 1, scaleResolutionDownTo: { maxWidth: 640, maxHeight: 360 } };

const camera = { kind: 'video', getSettings: () => ({ width: 1920, height: 1080 }) };
const share = { kind: 'video', getSettings: () => ({ displaySurface: 'monitor' }) };

const saved = {};

function load(sender) {
	saved.RTCRtpSender = globalThis.RTCRtpSender;
	saved.RTCPeerConnection = globalThis.RTCPeerConnection;
	globalThis.RTCRtpSender = { prototype: { setParameters: (p) => p } };
	globalThis.RTCPeerConnection = { prototype: { addTransceiver: (t, init) => init, setRemoteDescription: (d) => d } };
	delete require.cache[modulePath];
	require(modulePath).init({ media: { camera: { sender: { enabled: true, maxWidth: 1920, maxHeight: 1080, maxBitrate: 4000000, ...sender } } } });
}

function set(encodings, track = camera) {
	return globalThis.RTCRtpSender.prototype.setParameters.call({ track }, { encodings: structuredClone(encodings) }).encodings;
}

function lifted(enc, extra = {}) {
	return { ...enc, scaleResolutionDownTo: LIFTED, scaleResolutionDownBy: 1, maxBitrate: 4000000, ...extra };
}

beforeEach(() => {
	mock.method(console, 'debug', () => {});
	mock.method(console, 'info', () => {});
});

afterEach(() => {
	globalThis.RTCRtpSender = saved.RTCRtpSender;
	globalThis.RTCPeerConnection = saved.RTCPeerConnection;
	mock.restoreAll();
});

describe('cameraSendResolution, default liftFromHeight 720', () => {
	beforeEach(() => load({}));

	it('lifts the 720p main layer and keeps the thumbnail and unused layer', () => {
		assert.deepStrictEqual(set([THUMB, MAIN_720, UNUSED]), [THUMB, lifted(MAIN_720), UNUSED]);
	});

	it('leaves a 540p main layer alone', () => {
		assert.deepStrictEqual(set([THUMB, MAIN_540, UNUSED]), [THUMB, MAIN_540, UNUSED]);
	});

	it('lifts the only encoding of a 1:1 call and sets the frame rate cap', () => {
		const one = { active: true, maxBitrate: 1500000, scaleResolutionDownTo: { maxWidth: 1280, maxHeight: 720 } };
		assert.deepStrictEqual(set([one]), [lifted(one, { maxFramerate: 30 })]);
	});

	it('treats an encoding without an active key as active', () => {
		const one = { maxBitrate: 1500000, scaleResolutionDownBy: 1.5 };
		assert.deepStrictEqual(set([one]), [lifted(one, { maxFramerate: 30 })]);
	});

	it('reads the width and height form of scaleResolutionDownTo', () => {
		const out = set([{ active: true, scaleResolutionDownTo: { width: 1280, height: 720 } }]);
		assert.deepStrictEqual(out[0].scaleResolutionDownTo, LIFTED);
	});

	it('gives the other layers a downTo matching their downBy when none had one', () => {
		const out = set([
			{ rid: '1', active: true, maxBitrate: 133200, scaleResolutionDownBy: 6 },
			{ rid: '2', active: true, maxBitrate: 1000000, scaleResolutionDownBy: 1.5 },
			{ rid: '3', active: false, scaleResolutionDownBy: 1 },
		]);
		assert.deepStrictEqual(out[0], { rid: '1', active: true, maxBitrate: 133200, scaleResolutionDownBy: 6, scaleResolutionDownTo: { maxWidth: 320, maxHeight: 180 } });
		assert.deepStrictEqual(out[1].scaleResolutionDownTo, LIFTED);
		assert.deepStrictEqual(out[2].scaleResolutionDownTo, { maxWidth: 640, maxHeight: 360 });
	});

	it('rewrites addTransceiver sendEncodings the same way', () => {
		const init = globalThis.RTCPeerConnection.prototype.addTransceiver.call({}, camera, { sendEncodings: structuredClone([THUMB, MAIN_720, UNUSED]) });
		assert.deepStrictEqual(init.sendEncodings, [THUMB, lifted(MAIN_720), UNUSED]);
	});

	it('leaves screen sharing alone', () => {
		assert.deepStrictEqual(set([MAIN_720], share), [MAIN_720]);
	});
});

describe('cameraSendResolution, liftFromHeight 0', () => {
	beforeEach(() => load({ liftFromHeight: 0 }));

	it('lifts a 540p main layer', () => {
		assert.deepStrictEqual(set([THUMB, MAIN_540, UNUSED]), [THUMB, lifted(MAIN_540), UNUSED]);
	});

	it('lifts a 180p 15 fps main layer to 30 fps and keeps the thumbnail', () => {
		assert.deepStrictEqual(set([THUMB, MAIN_180, UNUSED]), [THUMB, lifted(MAIN_180, { maxFramerate: 30 }), UNUSED]);
	});

	it('never lifts the thumbnail when it is the only active layer', () => {
		const off = { ...MAIN_720, active: false };
		assert.deepStrictEqual(set([THUMB, off, UNUSED]), [THUMB, off, UNUSED]);
	});

	it('keeps a frame rate Teams set above the cap', () => {
		assert.strictEqual(set([THUMB, { ...MAIN_720, maxFramerate: 60 }, UNUSED])[1].maxFramerate, 60);
	});

	it('leaves every layer alone when all are inactive', () => {
		const all = [{ ...THUMB, active: false }, { ...MAIN_720, active: false }, UNUSED];
		assert.deepStrictEqual(set(all), all);
	});
});
