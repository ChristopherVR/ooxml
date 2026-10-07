// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { afterEach, expect, it, vi } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { loadDocx, type Paragraph } from 'ooxml-core/docx';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { wordYjsPluginKey } from 'ooxml-core/docx/ui';
import { DocxEditorElement } from './index';
import './index';
import { applyLink, followLinkAt, linkAtSelection } from './link-commands';

const editors: DocxEditorElement[] = [];
const sessions: CollabSession[] = [];
const viewOf = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
afterEach(() => {
	for (const editor of editors.splice(0)) {
		editor.stopCollaboration(true);
		editor.remove();
	}
	for (const session of sessions.splice(0)) session.destroy();
	vi.restoreAllMocks();
});

it('opens, retargets, removes and exports a peer picture hyperlink through the shared command and dialog', async () => {
	const bytes = new Uint8Array(
		await readFile(
			'../core/docx/__fixtures__/review-advanced-object-formatting/picture-before.docx',
		),
	);
	const hub = createMemoryHub();
	for (const name of ['Ada', 'Bob']) {
		sessions.push(
			createCollabSession({
				roomId: 'picture-links',
				provider: transportProvider({ transport: hub.createTransport('picture-links') }),
				user: { name },
				heartbeatMs: 0,
				teardown: false,
			}),
		);
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		editors.push(editor);
		await editor.load(bytes);
	}
	const [a, b] = editors;
	const av = viewOf(a!);
	let pos = -1;
	av.state.doc.descendants((node, position) => {
		if (node.type.name === 'image') pos = position;
	});
	av.dispatch(av.state.tr.setSelection(TextSelection.create(av.state.doc, pos, pos + 1)));
	applyLink(av, { href: 'https://example.com', tooltip: 'Example' });
	a!.startYjsCollaboration(sessions[0]!, { documentId: 'source', initializeIfEmpty: true });
	b!.startYjsCollaboration(sessions[1]!, { documentId: 'source' });
	const bv = viewOf(b!);
	bv.dispatch(bv.state.tr.setSelection(TextSelection.create(bv.state.doc, pos, pos + 1)));
	expect(linkAtSelection(bv)).toEqual({ href: 'https://example.com', tooltip: 'Example' });
	const open = vi.spyOn(window, 'open').mockReturnValue(null);
	expect(followLinkAt(bv, pos)).toBe(true);
	expect(open).toHaveBeenCalledWith('https://example.com', '_blank', 'noopener,noreferrer');
	const collab = wordYjsPluginKey.getState(bv.state)!;
	collab.stopCapturing();
	const shadow = b!.shadowRoot!;
	shadow.querySelector<HTMLButtonElement>('[aria-label="Insert link"]')!.click();
	const dialog = shadow.querySelector<HTMLElement>('.dve-link-dialog')!;
	const address = dialog.querySelector<HTMLInputElement>('[aria-label="Address"]')!;
	expect(address.value).toBe('https://example.com');
	address.value = 'https://example.org';
	dialog.querySelector<HTMLInputElement>('[aria-label="ScreenTip"]')!.value = 'Updated';
	const button = (label: string) =>
		[...dialog.querySelectorAll('button')].find((item) => item.textContent === label)!;
	button('Insert').click();
	expect(viewOf(a!).state.doc.eq(bv.state.doc)).toBe(true);
	for (const editor of editors) {
		const reloaded = await loadDocx(await editor.saveBytes());
		const picture = (reloaded.model.blocks[0] as Paragraph).runs.find((run) => run.image)!;
		expect(picture.link).toEqual({ href: 'https://example.org', tooltip: 'Updated' });
		expect(reloaded.media!.get(picture.image!.partName)).toBeDefined();
	}
	expect(collab.undo()).toBe(true);
	expect(linkAtSelection(bv)?.href).toBe('https://example.com');
	collab.stopCapturing();
	shadow.querySelector<HTMLButtonElement>('[aria-label="Insert link"]')!.click();
	button('Remove link').click();
	expect(viewOf(a!).state.doc.eq(bv.state.doc)).toBe(true);
	expect(linkAtSelection(bv)).toBeUndefined();
	expect(collab.undo()).toBe(true);
	expect(linkAtSelection(bv)?.href).toBe('https://example.com');
});
