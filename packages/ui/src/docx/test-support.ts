import type { Block, Paragraph, Table } from 'ooxml-core/docx';
import { expectDefined } from './defined';

/** The element at `index`, failing the test with a clear message when the list is too short. */
export function at<T>(list: ArrayLike<T>, index: number): T {
	return expectDefined(list[index], `item ${index} of ${list.length}`);
}

/** Asserts a value is present (not undefined/null) and returns it with a narrowed type. */
export function must<T>(value: T | null | undefined, what = 'value'): T {
	return expectDefined(value, what);
}

/** The paragraph block at `index`, failing the test when it is missing or is a table. */
export function paragraphAt(blocks: readonly Block[], index: number): Paragraph {
	const block = at(blocks, index);
	if (block.type !== 'paragraph') throw new Error(`Expected block ${index} to be a paragraph.`);
	return block;
}

/** The table block at `index`, failing the test when it is missing or is a paragraph. */
export function tableAt(blocks: readonly Block[], index: number): Table {
	const block = at(blocks, index);
	if (block.type !== 'table') throw new Error(`Expected block ${index} to be a table.`);
	return block;
}

/**
 * A button of the title bar's Quick Access strip ("Save", "Undo", "Redo"). The strip is drawn by
 * the shared `office-ui-title-bar`, inside its own shadow root, once that element has rendered.
 */
export async function quickAccessButton(
	editor: { shadowRoot: ShadowRoot | null },
	label: string,
): Promise<HTMLButtonElement> {
	const bar = must(
		editor.shadowRoot?.querySelector<HTMLElement & { updateComplete: Promise<unknown> }>(
			'office-ui-title-bar',
		),
		'title bar',
	);
	await bar.updateComplete;
	return must(
		bar.shadowRoot?.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`),
		`quick access button ${label}`,
	);
}

type Updating = HTMLElement & { updateComplete: Promise<unknown> };

/** The shadow root of the shared ribbon in `scope`, after it has rendered its tabs. */
export async function ribbonShadow(scope: ParentNode): Promise<ShadowRoot> {
	const ribbon =
		scope instanceof HTMLElement && scope.localName === 'office-ui-ribbon'
			? (scope as Updating)
			: must(scope.querySelector<Updating>('office-ui-ribbon'), 'ribbon');
	await ribbon.updateComplete;
	return must(ribbon.shadowRoot, 'ribbon shadow root');
}

/** The rendered button of ribbon tab `key` ("home", "insert", "table"...). */
export async function ribbonTab(scope: ParentNode, key: string): Promise<HTMLButtonElement> {
	const root = await ribbonShadow(scope);
	return must(
		root.querySelector<HTMLButtonElement>(`[role="tab"][data-tab="${key}"]`),
		`tab ${key}`,
	);
}

/** The tab buttons, in order. */
export async function ribbonTabs(scope: ParentNode): Promise<HTMLButtonElement[]> {
	return [...(await ribbonShadow(scope)).querySelectorAll<HTMLButtonElement>('[role="tab"]')];
}

/** The File button of the shared ribbon. */
export async function ribbonFile(scope: ParentNode): Promise<HTMLButtonElement> {
	return must((await ribbonShadow(scope)).querySelector<HTMLButtonElement>('.file'), 'File button');
}
