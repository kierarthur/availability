import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
const source = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8');
const context = vm.createContext({});
new vm.Script(source.slice(source.indexOf('function availabilitySaveFeedback('), source.indexOf('async function submitChanges('))).runInContext(context);
const feedback = bridge_sync => context.availabilitySaveFeedback({ results: [{ ymd: '2026-10-12', applied: true }], bridge_sync });
test('confirmed synchronization and legacy-only saves retain clear truthful feedback', () => {
  assert.match(feedback({ state: 'SYNCED' }).message, /sheet and MyTMS/);
  for (const state of ['SYNCED', 'LEGACY_ONLY', 'NO_CHANGES']) assert.equal(feedback({ state }).unconfirmed, false);
  assert.equal(feedback(undefined).message, 'Saved 1 change.'); // backwards-compatible older API
});
test('pending and unknown receipts preserve intent instead of falsely saying synchronized', () => {
  for (const state of ['PENDING', 'UNKNOWN']) {
    assert.equal(feedback({ state }).unconfirmed, true);
    assert.match(feedback({ state }).message, /sync is not confirmed/);
  }
  assert.match(source, /if \(!feedback.unconfirmed\) draft = \{\}/);
});
test('queued Sheet changes are not described as saved or synchronized', () => {
  assert.match(feedback({ state: 'QUEUED' }).message, /queued/);
  assert.equal(feedback({ state: 'QUEUED' }).unconfirmed, false);
});
test('canonical per-date rejection is explicit and unsafe references are not displayed', () => {
  assert.match(feedback({ state: 'REJECTED' }).message, /did not accept every date/);
  assert.equal(feedback({ state: 'PENDING', correlation_id: '<secret-phone>' }).message.includes('<secret-phone>'), false);
  assert.match(feedback({ state: 'PENDING', correlation_id: '01M4T000000000000000000001' }).message, /Reference: 01M4/);
});
test('the timestamp describes a load, not a worker availability update', () => {
  assert.match(source, /lastLoaded.textContent = 'Loaded '/);
});
