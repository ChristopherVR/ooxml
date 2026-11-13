import { expect, test, vi } from 'vitest';

import { BUILDING_BLOCK_OK } from './building-block-actions';
import type { BuildingBlockResult } from './building-block-actions';
import {
	appPropertiesChanged,
	buildDocumentPropertiesActions,
	documentPropertiesChanged,
	planDocumentPropertiesPatch,
} from './building-block-document-properties';
import type { DocumentPropertiesSnapshot } from './building-block-document-properties';
import {
	buildFontEmbeddingActions,
	describeFontEmbeddingState,
} from './building-block-font-embedding';
import {
	buildPasswordProtectionActions,
	describePasswordProtection,
	markAsFinalProperties,
	passwordStrength,
	validateOpenPassword,
} from './building-block-password';
import { buildSignatureActions, describeSignaturesState } from './building-block-signatures';

const snapshot = (): DocumentPropertiesSnapshot => ({
	core: { title: 'Deck', creator: 'Ada' },
	custom: [{ name: 'Project', value: 'X', type: 'lpwstr' }],
	app: { company: 'Acme', slides: 3 },
});

test('open passwords are validated in the bundled dialog order', () => {
	expect(validateOpenPassword('', '')).toEqual({ ok: false, code: 'password-required' });
	expect(validateOpenPassword('abcd', 'abce')).toEqual({ ok: false, code: 'password-mismatch' });
	expect(validateOpenPassword('abc', 'abc')).toEqual({ ok: false, code: 'password-too-short' });
	expect(validateOpenPassword('abcd', 'abcd')).toEqual(BUILDING_BLOCK_OK);
	expect(passwordStrength('')).toBe(0);
	expect(passwordStrength('Abcdefgh1!xyz')).toBe(4);
});

test('protection state reports flags and never a secret', () => {
	const state = describePasswordProtection({
		openPasswordSet: true,
		recommendation: { kind: 'modifyVerifier' },
		hasModifyVerifier: true,
		customProperties: [{ name: '_MarkAsFinal', value: 'true', type: 'bool' }],
	});
	expect(state).toMatchObject({
		openPassword: true,
		modifyPassword: true,
		readOnlyRecommended: true,
		markedFinal: true,
	});
	expect(state.supported.modifyPassword).toBe(false);
	expect(JSON.stringify(state)).not.toMatch(/hash|salt/i);
});

test('mark as final adds or removes the property and ignores a no-op', () => {
	const marked = markAsFinalProperties([], true);
	expect(marked).toEqual([{ name: '_MarkAsFinal', value: 'true', type: 'bool' }]);
	expect(markAsFinalProperties(marked ?? [], true)).toBeUndefined();
	expect(markAsFinalProperties(marked ?? [], false)).toEqual([]);
});

function passwordActions(refuse: () => BuildingBlockResult | undefined) {
	const applyOpenPassword = vi.fn();
	const markDirty = vi.fn();
	const applyCustomProperties = vi.fn();
	const actions = buildPasswordProtectionActions({
		state: describePasswordProtection({
			openPasswordSet: false,
			recommendation: { kind: null },
			hasModifyVerifier: false,
			customProperties: [],
		}),
		customProperties: () => [],
		refuse,
		applyOpenPassword,
		applyCustomProperties,
		markDirty,
		complete: (result) => result,
	});
	return { actions, applyOpenPassword, markDirty, applyCustomProperties };
}

test('setting and clearing the open password applies and marks the deck changed', () => {
	const { actions, applyOpenPassword, markDirty } = passwordActions(() => undefined);
	expect(actions.setOpenPassword('abc', 'abc')).toEqual({ ok: false, code: 'password-too-short' });
	expect(applyOpenPassword).not.toHaveBeenCalled();
	expect(actions.setOpenPassword('secret', 'secret')).toEqual(BUILDING_BLOCK_OK);
	expect(applyOpenPassword).toHaveBeenLastCalledWith('secret');
	expect(actions.clearOpenPassword()).toEqual(BUILDING_BLOCK_OK);
	expect(applyOpenPassword).toHaveBeenLastCalledWith(null);
	expect(markDirty).toHaveBeenCalledTimes(2);
	expect(actions.setModifyPassword('x')).toEqual({ ok: false, code: 'unsupported' });
});

test('password operations are refused when the deck is not editable', () => {
	const denied = { ok: false, code: 'read-only' } as const;
	const { actions, applyOpenPassword, applyCustomProperties } = passwordActions(() => denied);
	expect(actions.setOpenPassword('secret', 'secret')).toEqual(denied);
	expect(actions.clearOpenPassword()).toEqual(denied);
	expect(actions.setMarkedFinal(true)).toEqual(denied);
	expect(actions.setModifyPassword(null)).toEqual(denied);
	expect(applyOpenPassword).not.toHaveBeenCalled();
	expect(applyCustomProperties).not.toHaveBeenCalled();
});

test('document property patches change only what they name', () => {
	const plan = planDocumentPropertiesPatch(snapshot(), {
		core: { title: 'New' },
		app: { company: 'Acme' },
		custom: {
			set: [
				{ name: 'Project', value: 'Y', type: 'lpwstr' },
				{ name: 'N', value: '3', type: 'i4' },
			],
		},
	});
	expect(plan).toMatchObject({ ok: true, changed: true, core: { title: 'New' } });
	if (plan.ok) {
		expect(plan.app).toBeUndefined();
		expect(plan.custom).toEqual([
			{ name: 'Project', value: 'Y', type: 'lpwstr' },
			{ name: 'N', value: '3', type: 'i4' },
		]);
	}
	expect(planDocumentPropertiesPatch(snapshot(), { core: { title: 'Deck' } })).toMatchObject({
		ok: true,
		changed: false,
	});
});

test('malformed property patches are rejected as a whole', () => {
	const bad = [
		{ custom: { set: [{ name: ' ', value: '', type: 'lpwstr' }] } },
		{ custom: { set: [{ name: 'a', value: 'x', type: 'i4' }] } },
		{ custom: { set: [{ name: 'a', value: 'maybe', type: 'bool' }] } },
		{ custom: { set: [{ name: 'a', value: '', type: 'weird' }] } },
		{ core: { title: 5 as unknown as string } },
		{ core: { created: '2020' } as never },
	];
	for (const patch of bad) {
		expect(planDocumentPropertiesPatch(snapshot(), patch)).toEqual({
			ok: false,
			code: 'invalid-data',
		});
	}
});

test('applying properties hands the viewer only the changed groups and rechecks permission', () => {
	const applyUpdate = vi.fn();
	let refusal: BuildingBlockResult | undefined;
	const actions = buildDocumentPropertiesActions({
		current: snapshot,
		refuse: () => refusal,
		applyUpdate,
		complete: (result) => result,
	});
	expect(actions.apply({ core: { subject: 'S' } })).toEqual(BUILDING_BLOCK_OK);
	expect(applyUpdate).toHaveBeenCalledWith({ core: { subject: 'S' } });
	applyUpdate.mockClear();
	actions.apply({ core: { title: 'Deck' } });
	expect(applyUpdate).not.toHaveBeenCalled();
	refusal = { ok: false, code: 'read-only' };
	expect(actions.apply({ core: { subject: 'S' } })).toEqual(refusal);
	expect(applyUpdate).not.toHaveBeenCalled();
});

test('the dialog dirty check and the shared check agree', () => {
	const current = { core: { title: 'a' }, custom: [], app: { company: 'c' } };
	const same = { core: { title: 'a' }, custom: [], app: { company: 'c' } };
	expect(documentPropertiesChanged(current, same)).toBe(false);
	expect(documentPropertiesChanged(current, { core: { title: 'b' }, custom: [] })).toBe(true);
	expect(appPropertiesChanged(undefined, { manager: 'm' })).toBe(true);
});

function fontActions(embedFonts: boolean, families: string[]) {
	const state = describeFontEmbeddingState({
		embedFonts,
		embeddedFamilies: families,
		usedFamilies: ['Aptos'],
	});
	const apply = vi.fn();
	const markDirty = vi.fn();
	const actions = buildFontEmbeddingActions({
		state,
		refuse: () => undefined,
		latest: () => ({ embedFonts, canEmbed: state.canEmbed }),
		apply,
		markDirty,
		complete: (result) => result,
	});
	return { state, actions, apply, markDirty };
}

test('font embedding refuses to embed what the deck does not carry', () => {
	const none = fontActions(false, []);
	expect(none.state).toMatchObject({
		canEmbed: false,
		embeddedFamilies: [],
		usedFamilies: ['Aptos'],
	});
	expect(none.state.unavailableReasonKey).toBeDefined();
	expect(none.actions.setEmbedFonts(true)).toEqual({ ok: false, code: 'unsupported' });
	expect(none.apply).not.toHaveBeenCalled();
	const some = fontActions(true, ['Aptos', 'aptos']);
	expect(some.state.embeddedFamilies).toEqual(['Aptos']);
	expect(some.actions.setEmbedFonts(false)).toEqual(BUILDING_BLOCK_OK);
	expect(some.apply).toHaveBeenCalledWith(false);
	expect(some.markDirty).toHaveBeenCalledTimes(1);
	some.apply.mockClear();
	some.actions.setEmbedFonts(true);
	expect(some.apply).not.toHaveBeenCalled();
});

test('signatures can be listed and removed on save but never created', async () => {
	const markDirty = vi.fn();
	const make = (signed: boolean) =>
		buildSignatureActions({
			state: describeSignaturesState({ signed, count: signed ? 1 : 0 }),
			content: () => null,
			refuse: () => undefined,
			markDirty,
			complete: (result) => result,
		});
	const signed = make(true);
	expect(await signed.list()).toEqual([]);
	expect(signed.remove()).toEqual(BUILDING_BLOCK_OK);
	expect(markDirty).toHaveBeenCalledTimes(1);
	expect(signed.sign()).toEqual({ ok: false, code: 'unsupported' });
	expect(signed.state.supported).toMatchObject({ create: false, trust: false });
	expect(make(false).remove()).toEqual({ ok: false, code: 'no-signatures' });
});
