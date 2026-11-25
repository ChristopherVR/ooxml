// @vitest-environment jsdom
import { createDocument, type StyleEdit } from 'ooxml-core/docx';
import { afterEach, expect, it } from 'vitest';
import { isDialogOpen } from './dialog-shell';
import { createStyleDialog } from './style-dialog';

afterEach(() => document.body.replaceChildren());

it('opens the shared dialog, validates a new style and submits its typed fields', () => {
	const edits: StyleEdit[] = [];
	const dialog = createStyleDialog({
		model: createDocument(),
		locale: 'en',
		modify: false,
		selection: {
			runFormatting: { bold: true },
			paragraphFormatting: { align: 'center' },
			paragraphId: 'Normal',
			characterId: '',
			tableId: '',
		},
		submit: (edit) => Boolean(edits.push(edit)),
	});
	const root = dialog.element;
	document.body.append(root);
	dialog.open();
	const field = (label: string) => root.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
	const ok = [...root.querySelectorAll('button')].find((button) => button.textContent === 'OK')!;
	expect(isDialogOpen(root)).toBe(true);
	expect(field('Based on').value).toBe('Normal');
	expect(field('Bold').value).toBe('on');
	ok.click();
	expect(root.querySelector<HTMLElement>('[role="alert"]')?.hidden).toBe(false);
	expect(edits).toHaveLength(0);
	field('Style name').value = 'Callout';
	field('Style for following paragraph').value = 'Heading1';
	field('Font').value = 'Georgia';
	field('Font size').value = '12.5';
	field('After').value = '4.5';
	ok.click();
	expect(edits).toEqual([
		{
			kind: 'paragraph',
			name: 'Callout',
			basedOn: 'Normal',
			next: 'Heading1',
			runFormatting: { bold: true, fontFamily: 'Georgia', fontSize: 12.5 },
			paragraphFormatting: { align: 'center', spacingAfterTwips: 90 },
		},
	]);
	expect(root.isConnected).toBe(false);
});

it('chooses the current style to modify and cancels without submitting', () => {
	const dialog = createStyleDialog({
		model: createDocument(),
		locale: 'fr',
		modify: true,
		selection: {
			runFormatting: {},
			paragraphFormatting: {},
			paragraphId: 'Heading1',
			characterId: '',
			tableId: '',
		},
		submit: () => {
			throw new Error('Cancel submitted the style');
		},
	});
	const root = dialog.element;
	document.body.append(root);
	dialog.open();
	expect(root.querySelector<HTMLInputElement>('[aria-label="Nom du style"]')?.value).toBe(
		'heading 1',
	);
	expect(root.querySelector<HTMLSelectElement>('[aria-label="Type de style"]')?.disabled).toBe(
		true,
	);
	root.dispatchEvent(new CustomEvent('office-dialog-close', { cancelable: true }));
	expect(root.isConnected).toBe(false);
});
