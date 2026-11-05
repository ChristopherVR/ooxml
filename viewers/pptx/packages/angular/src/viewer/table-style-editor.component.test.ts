/**
 * table-style-editor.component.test.ts: the table STYLE DEFINITION editor
 * ("Edit style...") lets an author change a `ppt/tableStyles.xml` section's
 * fill/text/borders, and clone/delete a style, entirely through the shared
 * `pptx-viewer-shared` describe/apply pair.
 *
 * No TestBed (matching the rest of this package): the component is
 * constructed inside a plain `Injector` context, mirroring
 * `slide-background-card.component.test.ts`.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { OutputEmitterRef } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { ParsedTableStyleMap } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { LoadContentService } from './load-content.service';
import { TableStyleEditorComponent } from './table-style-editor.component';

const STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

function styleMap(): ParsedTableStyleMap {
	return {
		[STYLE_ID]: {
			styleId: STYLE_ID,
			styleName: 'Medium Style 2 - Accent 1',
			wholeTblFill: { schemeColor: '', color: '#336699' },
		},
	};
}

function harness(map: ParsedTableStyleMap | undefined, id: string | undefined, canEdit = true) {
	const injector = Injector.create({
		providers: [
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			{ provide: LoadContentService, useValue: { themeColorMap: () => undefined } },
		],
	});
	const component = runInInjectionContext(injector, () => new TableStyleEditorComponent());
	(component as unknown as { styleMap: () => ParsedTableStyleMap | undefined }).styleMap = () =>
		map;
	(component as unknown as { styleId: () => string | undefined }).styleId = () => id;
	(component as unknown as { canEdit: () => boolean }).canEdit = () => canEdit;
	return component;
}

/** Answer the in-app dialog the component opened: type into its input, then press a button. */
async function answerDialog(button: 'confirm' | 'cancel', value?: string): Promise<void> {
	const dialog = await vi.waitFor(() => {
		const found = document.querySelector('office-ui-message-dialog');
		if (!found) throw new Error('no dialog yet');
		return found;
	});
	const input = dialog.shadowRoot?.querySelector('input');
	if (input && value !== undefined) input.value = value;
	(dialog.shadowRoot!.querySelector(`button.${button}`) as HTMLButtonElement).click();
	await new Promise((resolve) => setTimeout(resolve, 0));
}

afterEach(() => {
	document.querySelectorAll('office-ui-message-dialog').forEach((node) => node.remove());
});

describe('tableStyleEditorComponent', () => {
	it('has no entry without a styleId', () => {
		const component = harness(styleMap(), undefined);
		expect(component['entry']()).toBeUndefined();
	});

	it('resolves the entry for the assigned styleId', () => {
		const component = harness(styleMap(), STYLE_ID);
		expect(component['entry']()?.styleId).toBe(STYLE_ID);
	});

	it('applies a fill-colour edit and emits the updated map', () => {
		const component = harness(styleMap(), STYLE_ID);
		let emitted: ParsedTableStyleMap | undefined;
		vi.spyOn(
			component.styleMapChange as OutputEmitterRef<ParsedTableStyleMap>,
			'emit',
		).mockImplementation((value) => {
			emitted = value;
		});
		component['onFieldEdit']({ kind: 'fillColor', hex: '#ff0000', ref: undefined });
		expect(emitted?.[STYLE_ID].wholeTblFill).toStrictEqual({ schemeColor: '', color: '#ff0000' });
	});

	it('creates a new style from the current one and assigns it', async () => {
		const component = harness(styleMap(), STYLE_ID);
		let emittedMap: ParsedTableStyleMap | undefined;
		let assignedId: string | undefined;
		vi.spyOn(
			component.styleMapChange as OutputEmitterRef<ParsedTableStyleMap>,
			'emit',
		).mockImplementation((value) => {
			emittedMap = value;
		});
		vi.spyOn(component.assignStyle as OutputEmitterRef<string>, 'emit').mockImplementation(
			(value) => {
				assignedId = value;
			},
		);
		const creating = component['createFromCurrent']();
		await answerDialog('confirm', 'My Custom Style');
		await creating;
		expect(Object.keys(emittedMap ?? {})).toHaveLength(2);
		const newId = Object.keys(emittedMap ?? {}).find((id) => id !== STYLE_ID);
		expect(emittedMap?.[newId as string].styleName).toBe('My Custom Style');
		expect(assignedId).toBe(newId);
	});

	it('creates a brand-new style with no current style selected and assigns it', async () => {
		const component = harness(undefined, undefined);
		let emittedMap: ParsedTableStyleMap | undefined;
		let assignedId: string | undefined;
		vi.spyOn(
			component.styleMapChange as OutputEmitterRef<ParsedTableStyleMap>,
			'emit',
		).mockImplementation((value) => {
			emittedMap = value;
		});
		vi.spyOn(component.assignStyle as OutputEmitterRef<string>, 'emit').mockImplementation(
			(value) => {
				assignedId = value;
			},
		);
		const creating = component['createFromCurrent']();
		await answerDialog('confirm', 'Brand New Style');
		await creating;
		const ids = Object.keys(emittedMap ?? {});
		expect(ids).toHaveLength(1);
		expect(emittedMap?.[ids[0]].styleName).toBe('Brand New Style');
		expect(assignedId).toBe(ids[0]);
	});

	it('deletes the style after confirmation', async () => {
		const component = harness(styleMap(), STYLE_ID);
		let emittedMap: ParsedTableStyleMap | undefined;
		let deletedId: string | undefined;
		let closed = false;
		vi.spyOn(
			component.styleMapChange as OutputEmitterRef<ParsedTableStyleMap>,
			'emit',
		).mockImplementation((value) => {
			emittedMap = value;
		});
		vi.spyOn(component.deleteStyle as OutputEmitterRef<string>, 'emit').mockImplementation(
			(value) => {
				deletedId = value;
			},
		);
		vi.spyOn(component.close as OutputEmitterRef<void>, 'emit').mockImplementation(() => {
			closed = true;
		});
		const deleting = component['handleDelete']();
		await answerDialog('confirm');
		await deleting;
		expect(Object.keys(emittedMap ?? {})).toHaveLength(0);
		expect(deletedId).toBe(STYLE_ID);
		expect(closed).toBeTruthy();
	});

	it('does not delete without confirmation', async () => {
		const component = harness(styleMap(), STYLE_ID);
		const emitSpy = vi.spyOn(
			component.styleMapChange as OutputEmitterRef<ParsedTableStyleMap>,
			'emit',
		);
		const deleting = component['handleDelete']();
		await answerDialog('cancel');
		await deleting;
		expect(emitSpy).not.toHaveBeenCalled();
	});
});
