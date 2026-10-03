import { afterEach, describe, expect, it } from 'vitest';
import { createBackstage } from './backstage.js';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import { registerViewerControls } from './office-ui.js';
import { ViewerShare } from './viewer-share.js';
import type { CancellableEditor } from './worker-editor.js';

const disposers: (() => void)[] = [];
afterEach(() => {
	disposers.splice(0).forEach((dispose) => dispose());
	document.body.replaceChildren();
});

const until = async (check: () => boolean) => {
	for (let i = 0; i < 100 && !check(); i++) await new Promise((r) => setTimeout(r, 10));
	expect(check()).toBe(true);
};

function window(name: string) {
	registerViewerControls();
	const parsed: number[][] = [];
	const editor: CancellableEditor = async (bytes) => ({
		bytes: new Uint8Array([...bytes, 7]),
		document: structuredClone(demoDocument),
		changedParts: ['visio/pages/page1.xml'],
		diagnostics: [],
	});
	const controller = new ViewerController(
		async (bytes) => {
			parsed.push([...new Uint8Array(bytes)]);
			return structuredClone(demoDocument);
		},
		() => {},
		editor,
	);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.append(createBackstage(document));
	const share = new ViewerShare(root, controller, () => ({
		displayName: name,
		avatarColor: '#16a34a',
	}));
	const dispose = share.wire();
	disposers.push(() => {
		dispose();
		controller.destroy();
	});
	const people = () =>
		(
			root.querySelector('office-ui-presence') as HTMLElement & {
				participants: { name: string }[];
			}
		).participants.map((person) => person.name);
	return { controller, share, root, parsed, people };
}

describe('File > Share', () => {
	it('rejects session names the core would refuse', async () => {
		const a = window('Ada');
		await a.share.start('no spaces');
		expect(a.share.active).toBe(false);
		expect(a.root.querySelector('.share-status')!.textContent).toMatch(/letters, digits/);
	});

	it('shares the drawing with another window and applies its edits as history steps', async () => {
		const room = `test-${Math.random().toString(36).slice(2, 8)}`;
		const a = window('Ada');
		await a.controller.load(new Uint8Array([1]));
		await a.share.start(room);
		expect(a.share.active).toBe(true);
		expect(a.root.querySelector<HTMLElement>('[data-share="stop"]')!.hidden).toBe(false);

		const b = window('Grace');
		await b.share.start(room);
		// The joiner had nothing open: it loads the room's drawing.
		await until(() => b.parsed.some((bytes) => bytes.join() === '1'));
		await until(() => b.people().includes('Ada') && a.people().includes('Grace'));

		await a.controller.applyEdits([]);
		await until(() => b.parsed.some((bytes) => bytes.join() === '1,7'));
		await until(() => !b.controller.state.edit.busy && b.controller.state.edit.canUndo);
		expect(b.controller.exportVsdx().bytes).toEqual(new Uint8Array([1, 7]));

		// The remote step is undoable, and the undo is shared back.
		await b.controller.undo();
		await until(() => a.parsed.some((bytes) => bytes.join() === '1'));
		b.share.stop();
		expect(b.people()).toEqual([]);
	});
});
