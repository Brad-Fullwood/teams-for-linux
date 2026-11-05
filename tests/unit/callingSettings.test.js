'use strict';

const { describe, it, beforeEach, afterEach, mock } = require('node:test');
const assert = require('node:assert');

const modulePath = require.resolve('../../app/browser/tools/callingSettings');

const OVERRIDES = {
	webrtcVideoCapabilityMaxFS: 8160,
	sendVideoCapabilityManager: { defaultFS: 8160 },
	video1080p: { software: { isEnabled: true } },
};

// Shape of the Teams calling config provider as far as the tool relies on it
class FakeProvider {
	constructor() {
		this._consoleOverrides = {};
		this.applied = 0;
	}

	setConsoleOverrides(e) {
		this._consoleOverrides = { ...this._consoleOverrides, ...e };
		this.applied++;
	}
}

function load(callingSettings) {
	delete require.cache[modulePath];
	require(modulePath).init({ media: { callingSettings } });
}

const tick = () => new Promise((r) => setTimeout(r, 5));

beforeEach(() => {
	for (const level of ['debug', 'info', 'warn', 'error']) mock.method(console, level, () => {});
});

afterEach(() => {
	delete Object.prototype._consoleOverrides;
	mock.restoreAll();
});

describe('callingSettings', () => {
	it('does nothing when no overrides are configured', () => {
		load(undefined);
		assert.strictEqual(Object.hasOwn(Object.prototype, '_consoleOverrides'), false);
		load({});
		assert.strictEqual(Object.hasOwn(Object.prototype, '_consoleOverrides'), false);
	});

	it('puts the overrides on the provider and asks Teams to apply them', async () => {
		load(OVERRIDES);
		const p = new FakeProvider();
		assert.deepStrictEqual(p._consoleOverrides, OVERRIDES);
		assert.strictEqual(Object.hasOwn(Object.prototype, '_consoleOverrides'), false);
		await tick();
		assert.strictEqual(p.applied, 1);
		assert.deepStrictEqual(p._consoleOverrides, OVERRIDES);
	});

	it('leaves later providers and plain objects alone', async () => {
		load(OVERRIDES);
		const first = new FakeProvider();
		const second = new FakeProvider();
		assert.deepStrictEqual(second._consoleOverrides, {});
		const plain = {};
		plain._consoleOverrides = 1;
		assert.strictEqual(plain._consoleOverrides, 1);
		await tick();
		assert.strictEqual(first.applied, 1);
		assert.strictEqual(second.applied, 0);
	});

	it('copies the overrides so Teams cannot change the configured object', () => {
		load(OVERRIDES);
		const p = new FakeProvider();
		p._consoleOverrides.sendVideoCapabilityManager.defaultFS = 1;
		assert.strictEqual(OVERRIDES.sendVideoCapabilityManager.defaultFS, 8160);
	});
});
