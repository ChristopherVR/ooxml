/**
 * The Share and Broadcast dialog rows name their input by their label, so a screen reader
 * announces "Server URL" for the server box and a test (or a user's tooling) finds it by name.
 */
import { describe, expect, it } from 'vitest';

import type { Translator } from '../../i18n';
import { createCopyField, createTextField } from './dialog-fields';

const t = ((key: string) => key) as Translator;

function labelledInput(row: HTMLElement): HTMLInputElement | null {
	const label = row.querySelector('label');
	const id = label?.getAttribute('for');
	return id ? (row.ownerDocument.getElementById(id) as HTMLInputElement | null) : null;
}

describe('dialog field rows', () => {
	it('tie each label to its own input', () => {
		const server = createTextField(document, 'Server URL', 'ws://', () => {});
		const room = createTextField(document, 'Room ID', 'room', () => {});
		const link = createCopyField(document, t, 'Viewer Link', () => {});
		document.body.append(server.el, room.el, link.el);
		try {
			expect(labelledInput(server.el)).toBe(server.input);
			expect(labelledInput(room.el)).toBe(room.input);
			expect(labelledInput(link.el)).toBe(link.input);
			expect(new Set([server.input.id, room.input.id, link.input.id]).size).toBe(3);
		} finally {
			server.el.remove();
			room.el.remove();
			link.el.remove();
		}
	});
});
