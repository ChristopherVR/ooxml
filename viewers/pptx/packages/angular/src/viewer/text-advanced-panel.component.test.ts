/**
 * TextAdvancedPanelComponent's Text Box options, Angular binding: the shared
 * `TEXT_BOX_SECTIONS` drawn by `pptx-inspector-fields`, merged into the
 * element's text style.
 *
 * No Angular TestBed (see `vitest.config.ts`): the component is instantiated
 * directly, inputs are stubbed as signals.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { componentSource } from './component-source.test-support';
import { TextAdvancedPanelComponent } from './text-advanced-panel.component';

describe('textAdvancedPanelComponent text box options', () => {
	it('mounts the shared Text Box descriptors', () => {
		const source = componentSource(import.meta.dirname, 'text-advanced-panel.component.ts');
		expect(source).toContain('[sections]="textBoxSections"');
		expect(source).not.toContain('txadv-valign');
	});

	it('reads the text style and merges a body-level patch into it', () => {
		const panel = runInInjectionContext(
			Injector.create({ providers: [] }),
			() => new TextAdvancedPanelComponent(),
		);
		const element = {
			type: 'text',
			id: 't1',
			x: 0,
			y: 0,
			width: 100,
			height: 40,
			textStyle: { fontSize: 18, vAlign: 'bottom' },
		} as PptxElement;
		Object.assign(panel, { element: signal(element) as unknown as InputSignal<PptxElement> });
		let emitted: Partial<PptxElement> | undefined;
		vi.spyOn(panel.patch as OutputEmitterRef<Partial<PptxElement>>, 'emit').mockImplementation(
			(value) => {
				emitted = value;
			},
		);
		expect(panel['textStyle']()).toStrictEqual({ fontSize: 18, vAlign: 'bottom' });
		panel['onTextBoxPatch']({ columnCount: 2 });
		expect((emitted as { textStyle?: unknown }).textStyle).toStrictEqual({
			fontSize: 18,
			vAlign: 'bottom',
			columnCount: 2,
		});
	});
});
