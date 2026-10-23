/**
 * The insert dialogs host maps the shared Insert SmartArt / Insert Chart dialog
 * events onto the editor. No Angular TestBed (see `vitest.config.ts`): the
 * component is instantiated directly with a stubbed editor.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { of } from 'rxjs';
import { describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { InsertDialogsHostComponent } from './insert-dialogs-host.component';

function create() {
	const addElement = vi.fn();
	const host = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: EditorStateService, useValue: { addElement } },
				{
					provide: TranslateService,
					useValue: {
						instant: (key: string) => (key === 'pptx.insertChart.title' ? 'Diagramm' : key),
						onLangChange: of(),
						onTranslationChange: of(),
					},
				},
			],
		}),
		() => new InsertDialogsHostComponent(),
	);
	Object.assign(host, { slideIndex: signal(2) as unknown as InputSignal<number> });
	return { host, addElement };
}
const call = (host: InsertDialogsHostComponent, name: string, detail: unknown) =>
	(host as unknown as Record<string, (event: Event) => void>)[name]!(
		new CustomEvent('x', { detail }),
	);

describe('insertDialogsHostComponent', () => {
	it('inserts the SmartArt layout and node texts the dialog chose', () => {
		const { host, addElement } = create();
		call(host, 'onInsertSmartArt', { layout: 'hierarchy', items: ['Boss', 'A', 'B'] });
		const [slide, element] = addElement.mock.calls[0]!;
		expect(slide).toBe(2);
		expect(element.type).toBe('smartArt');
		expect(element.smartArtData.layout).toBe('hierarchy');
		expect(element.smartArtData.nodes.map((node: { text: string }) => node.text)).toStrictEqual([
			'Boss',
			'A',
			'B',
		]);
	});

	it('inserts the chart type with the dialog title and grouping', () => {
		const { host, addElement } = create();
		call(host, 'onInsertChart', { kind: 'bar', title: 'Sales', grouping: 'stacked' });
		const [slide, element] = addElement.mock.calls[0]!;
		expect(slide).toBe(2);
		expect(element).toMatchObject({ type: 'chart', id: '' });
		expect(element.chartData).toMatchObject({
			chartType: 'bar',
			barDirection: 'bar',
			title: 'Sales',
			grouping: 'stacked',
		});
	});

	it('hands the dialogs a translator over ngx-translate', () => {
		const { host } = create();
		const translate = (host as unknown as { translate: () => (key: string) => string }).translate();
		expect(translate('pptx.insertChart.title')).toBe('Diagramm');
	});
});
