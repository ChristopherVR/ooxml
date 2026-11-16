/**
 * selection-pane.component.test.ts: the selection pane over the shared selection pane controller.
 *
 * This package has no TestBed (see `vitest.config.ts`), so the component is instantiated in an
 * injection context and driven through its handlers, and the static template contract (the e2e
 * data attributes, the rename input's events and accessible name) is read off the source as text.
 * The rename decision and the labels themselves are proved once in `ooxml-ui/pptx`.
 */
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { DestroyRef, Injector, runInInjectionContext, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { PptxElement } from 'pptx-viewer-core';
import { of } from 'rxjs';
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';
import { elementIcon } from './selection-pane-helpers';
import { SelectionPaneComponent } from './selection-pane.component';

const here = dirname(fileURLToPath(import.meta.url));
const source = componentSource(here, 'selection-pane.component.ts');

function element(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'shape',
		id: 'shape-1',
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		...overrides,
	} as PptxElement;
}

type Internals = {
	rows(): { element: PptxElement; index: number; name: string }[];
	startRename(id: string, name: string): void;
	setRenameDraft(event: Event): void;
	pane: { commitRename(): boolean; dragStart(i: number): boolean; drop(i: number): void };
};

function createPane(elements: PptxElement[], canEdit = true) {
	const translate = {
		instant: (key: string, params?: Record<string, unknown>) =>
			key === 'pptx.selectionPane.elementLabel'
				? `${String(params?.['type'])} ${String(params?.['number'])}`
				: key === 'pptx.elementType.shape'
					? 'Shape'
					: key,
		onLangChange: of(),
		onTranslationChange: of(),
		onDefaultLangChange: of(),
		onFallbackLangChange: of(),
	} as unknown as TranslateService;
	const pane = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: TranslateService, useValue: translate },
				{ provide: DestroyRef, useValue: { onDestroy: () => () => undefined } },
			],
		}),
		() => new SelectionPaneComponent(),
	);
	Object.defineProperty(pane, 'elements', { value: signal(elements) });
	Object.defineProperty(pane, 'canEdit', { value: signal(canEdit) });
	const renames: { id: string; name: string }[] = [];
	const moves: { from: number; to: number }[] = [];
	pane.renameElement.subscribe((event) => renames.push(event));
	pane.reorderElement.subscribe((event) => moves.push(event));
	return { pane, internals: pane as unknown as Internals, renames, moves };
}

const input = (value: string): Event => {
	const field = document.createElement('input');
	field.value = value;
	const event = new Event('input');
	Object.defineProperty(event, 'target', { value: field });
	return event;
};

describe('selection pane', () => {
	it('labels rows top-most first with the shared translated label', () => {
		const { internals } = createPane([element({ id: 'a' }), element({ id: 'b', name: 'Logo' })]);
		expect(internals.rows().map((row) => row.name)).toStrictEqual(['Logo', 'Shape 1']);
	});

	it('renames only when the box was edited, and clears with an empty name', () => {
		const { internals, renames } = createPane([element()]);
		internals.startRename('shape-1', 'Shape 1');
		internals.pane.commitRename();
		expect(renames).toStrictEqual([]);
		internals.startRename('shape-1', 'Shape 1');
		internals.setRenameDraft(input('Hero'));
		internals.pane.commitRename();
		internals.startRename('shape-1', 'Hero');
		internals.setRenameDraft(input(''));
		internals.pane.commitRename();
		expect(renames).toStrictEqual([
			{ id: 'shape-1', name: 'Hero' },
			{ id: 'shape-1', name: '' },
		]);
	});

	it('reorders by dragging, only on an editable deck', () => {
		const editable = createPane([element({ id: 'a' }), element({ id: 'b' })]);
		editable.internals.pane.dragStart(1);
		editable.internals.pane.drop(0);
		expect(editable.moves).toStrictEqual([{ from: 1, to: 0 }]);
		const readOnly = createPane([element({ id: 'a' }), element({ id: 'b' })], false);
		readOnly.internals.pane.dragStart(1);
		readOnly.internals.pane.drop(0);
		readOnly.internals.startRename('a', 'Shape 1');
		readOnly.internals.pane.commitRename();
		expect(readOnly.moves).toStrictEqual([]);
		expect(readOnly.renames).toStrictEqual([]);
	});

	it('falls back to "?" for an unknown element type icon', () => {
		expect(elementIcon('shape')).not.toBe('?');
		expect(elementIcon('somethingNew')).toBe('?');
	});
});

describe('selection pane template contract (source guards)', () => {
	it('stamps the e2e attributes on the pane root and each row name label', () => {
		expect(source).toContain('data-pptx-selection-pane');
		expect(source).toContain('data-pptx-selection-name');
	});

	it('renders a plain text input (role textbox) named from the shared i18n key', () => {
		expect(source).toContain('type="text"');
		expect(source).toContain(
			'[attr.aria-label]="\'pptx.selectionPane.renameElement\' | translate"',
		);
	});

	it('commits on Enter and blur, cancels on Escape', () => {
		expect(source).toContain('(keydown.enter)="pane.commitRename()"');
		expect(source).toContain('(blur)="pane.commitRename()"');
		expect(source).toContain('(keydown.escape)="cancelRename($event)"');
	});
});
