// @vitest-environment happy-dom
/**
 * The file-level dialogs of `useViewerBuildingBlocks` as host operations (#285): password
 * protection, document properties, font embedding and digital signatures. Each toolbar control
 * raises a request, `actions.<group>` reports the state and applies, and permission is rechecked
 * at execution.
 */
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, expect, test, vi } from 'vitest';

import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type {
	BuildingBlockRequest,
	UseViewerBuildingBlocksInput,
	ViewerBuildingBlocksResult,
} from './useViewerBuildingBlocks';

let root: Root;
let container: HTMLDivElement;
let latest: ViewerBuildingBlocksResult;
(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;
const onDirtyChange = vi.fn();
const onRequest = vi.fn<(request: BuildingBlockRequest) => void>();

function Harness({
	content,
	options,
}: {
	content: Uint8Array;
	options: Partial<UseViewerBuildingBlocksInput>;
}) {
	latest = useViewerBuildingBlocks({
		content,
		canEdit: true,
		autosaveEnabled: false,
		onDirtyChange,
		onRequest,
		...options,
	});
	return <div />;
}

async function mount(options: Partial<UseViewerBuildingBlocksInput> = {}) {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].isDirty = true;
	const content = await handler.save(data.slides);
	handler.dispose();
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	await act(async () => root.render(<Harness content={content} options={options} />));
	const deadline = Date.now() + 10000;
	while (latest.loading && Date.now() < deadline) {
		await act(async () => {
			await new Promise<void>((resolve) => {
				setTimeout(resolve, 0);
			});
		});
	}
	expect(latest.error).toBeNull();
}

afterEach(async () => {
	await act(async () => root?.unmount());
	container?.remove();
	onDirtyChange.mockClear();
	onRequest.mockClear();
	vi.unstubAllGlobals();
});

test('the four controls raise host requests instead of opening dialogs nothing renders', async () => {
	await mount();
	const cases: Array<[() => void, BuildingBlockRequest]> = [
		[() => latest.toolbarProps.onOpenPasswordProtection?.(), { kind: 'passwordProtection' }],
		[() => latest.toolbarProps.onOpenDocumentProperties?.(), { kind: 'documentProperties' }],
		[() => latest.toolbarProps.onOpenFontEmbedding?.(), { kind: 'fontEmbedding' }],
		[() => latest.toolbarProps.onOpenDigitalSignatures?.(), { kind: 'digitalSignatures' }],
	];
	for (const [open, expected] of cases) {
		await act(async () => open());
		expect(onRequest).toHaveBeenLastCalledWith(expected);
		expect(latest.actions.pending).toStrictEqual(expected);
	}
	for (const name of [
		'onOpenPasswordProtection',
		'onOpenDocumentProperties',
		'onOpenFontEmbedding',
		'onOpenDigitalSignatures',
	]) {
		expect(latest.actions.unmountedDialogActions).not.toContain(name);
	}
	for (const id of [
		'passwordProtection',
		'documentProperties',
		'fontEmbedding',
		'digitalSignatures',
	] as const) {
		expect(latest.actions.availability[id]).toMatchObject({ available: true, needsHostUi: true });
	}
});

test('a read-only viewer still shows the state, and refuses every change', async () => {
	await mount({ canEdit: false });
	await act(async () => latest.toolbarProps.onOpenDocumentProperties?.());
	expect(latest.actions.pending).toStrictEqual({ kind: 'documentProperties' });
	expect(latest.actions.availability.documentProperties).toMatchObject({
		available: false,
		reason: 'read-only',
	});
	const denied = { ok: false, code: 'read-only' };
	expect(latest.actions.documentProperties.apply({ core: { title: 'X' } })).toStrictEqual(denied);
	expect(latest.actions.passwordProtection.setOpenPassword('secret', 'secret')).toStrictEqual(
		denied,
	);
	expect(latest.actions.passwordProtection.setMarkedFinal(true)).toStrictEqual(denied);
	expect(latest.actions.fontEmbedding.setEmbedFonts(false)).toStrictEqual(denied);
	expect(latest.actions.digitalSignatures.remove()).toStrictEqual(denied);
	expect(onDirtyChange).not.toHaveBeenCalledWith(true);
});

test('open password: validated, applied, reported as state only, and marks the deck changed', async () => {
	await mount();
	expect(latest.actions.passwordProtection.state).toMatchObject({
		openPassword: false,
		modifyPassword: false,
		readOnlyRecommended: false,
		markedFinal: false,
	});
	await act(async () => latest.toolbarProps.onOpenPasswordProtection?.());
	expect(latest.actions.passwordProtection.setOpenPassword('abc', 'abc')).toStrictEqual({
		ok: false,
		code: 'password-too-short',
	});
	expect(latest.actions.passwordProtection.setOpenPassword('abcd', 'abce')).toStrictEqual({
		ok: false,
		code: 'password-mismatch',
	});
	let result: unknown;
	await act(async () => {
		result = latest.actions.passwordProtection.setOpenPassword('secret', 'secret');
	});
	expect(result).toStrictEqual({ ok: true });
	expect(latest.actions.pending).toBeNull();
	expect(latest.actions.passwordProtection.state.openPassword).toBeTruthy();
	expect(JSON.stringify(latest.actions.passwordProtection.state)).not.toContain('secret');
	expect(onDirtyChange).toHaveBeenCalledWith(true);
	await act(async () => {
		result = latest.actions.passwordProtection.clearOpenPassword();
	});
	expect(result).toStrictEqual({ ok: true });
	expect(latest.actions.passwordProtection.state.openPassword).toBeFalsy();
});

test('mark as final writes the custom property; a modify password is honestly unsupported', async () => {
	await mount();
	await act(async () => {
		latest.actions.passwordProtection.setMarkedFinal(true);
	});
	expect(latest.actions.passwordProtection.state.markedFinal).toBeTruthy();
	expect(latest.actions.documentProperties.current.custom).toContainEqual({
		name: '_MarkAsFinal',
		value: 'true',
		type: 'bool',
	});
	expect(latest.actions.passwordProtection.setModifyPassword('x')).toStrictEqual({
		ok: false,
		code: 'unsupported',
	});
	expect(latest.actions.passwordProtection.state.supported.modifyPassword).toBeFalsy();
});

test('document properties: patch core and custom fields, reject malformed ones', async () => {
	await mount();
	let result: unknown;
	await act(async () => {
		result = latest.actions.documentProperties.apply({
			core: { title: 'Host title', keywords: 'a,b' },
			custom: { set: [{ name: 'Reviewer', value: 'Ada', type: 'lpwstr' }] },
		});
	});
	expect(result).toStrictEqual({ ok: true });
	expect(latest.actions.documentProperties.current.core).toMatchObject({
		title: 'Host title',
		keywords: 'a,b',
	});
	expect(latest.actions.documentProperties.current.custom).toContainEqual({
		name: 'Reviewer',
		value: 'Ada',
		type: 'lpwstr',
	});
	expect(onDirtyChange).toHaveBeenCalledWith(true);
	expect(
		latest.actions.documentProperties.apply({
			custom: { set: [{ name: 'Count', value: 'many', type: 'i4' }] },
		}),
	).toStrictEqual({ ok: false, code: 'invalid-data' });
	await act(async () => {
		latest.actions.documentProperties.apply({ custom: { remove: ['Reviewer'] } });
	});
	expect(latest.actions.documentProperties.current.custom.map((p) => p.name)).not.toContain(
		'Reviewer',
	);
});

test('font embedding reports options and refuses to embed fonts the deck does not carry', async () => {
	await mount();
	const { state } = latest.actions.fontEmbedding;
	expect(state).toMatchObject({ canEmbed: false, embeddedFamilies: [] });
	expect(Array.isArray(state.usedFamilies)).toBeTruthy();
	expect(latest.actions.fontEmbedding.setEmbedFonts(true)).toStrictEqual({
		ok: false,
		code: 'unsupported',
	});
	await act(async () => {
		latest.actions.fontEmbedding.setEmbedFonts(false);
	});
	expect(latest.actions.fontEmbedding.state.embedFonts).toBeFalsy();
});

test('digital signatures: an unsigned deck lists none, has nothing to remove and cannot be signed', async () => {
	await mount();
	expect(latest.actions.digitalSignatures.state).toMatchObject({ signed: false, count: 0 });
	await expect(latest.actions.digitalSignatures.list()).resolves.toStrictEqual([]);
	expect(latest.actions.digitalSignatures.remove()).toStrictEqual({
		ok: false,
		code: 'no-signatures',
	});
	expect(latest.actions.digitalSignatures.sign()).toStrictEqual({
		ok: false,
		code: 'unsupported',
	});
	expect(latest.actions.digitalSignatures.state.supported).toMatchObject({
		create: false,
		trust: false,
	});
});
