/**
 * The inspector's Properties tab as a selection-aware Format pane (Angular):
 * the shared section list for the selection, open state kept per category in
 * the viewer's Format pane controller, launcher requests that expand and focus
 * a section, and the multi-selection rules.
 */
import { Component, CUSTOM_ELEMENTS_SCHEMA, signal } from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { componentSource } from './component-source.test-support';
import {
	FormatSectionsComponent,
	FormatSectionTemplateDirective,
} from './format-sections.component';
import { IsMobileService } from './is-mobile';
import { declareSignalFields } from './signal-inputs.test-support';
import { ViewerInspectorPanelService } from './viewer-inspector-panel.service';

type SectionHost = HTMLElement & { open: boolean; label: string; mixed: boolean };

const shape = (id: string, fillColor = '#ff0000'): PptxElement =>
	({
		id,
		type: 'shape',
		shapeType: 'rect',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeStyle: { fillColor },
		textStyle: {},
	}) as PptxElement;

@Component({
	standalone: true,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	imports: [FormatSectionsComponent, FormatSectionTemplateDirective],
	template: `
		<pptx-format-sections [elements]="elements()">
			<ng-template pptxFormatSection="fill"><div class="fill-body"></div></ng-template>
			<ng-template pptxFormatSection="sizePosition"><div class="size-body"></div></ng-template>
		</pptx-format-sections>
	`,
})
class HostComponent {
	readonly elements = signal<readonly PptxElement[]>([shape('a')]);
}

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
	declareSignalFields(HostComponent, import.meta.dirname);
});
afterEach(() => TestBed.resetTestingModule());

function open(): {
	fixture: ComponentFixture<HostComponent>;
	inspector: ViewerInspectorPanelService;
} {
	TestBed.configureTestingModule({
		imports: [HostComponent],
		providers: [
			provideTranslateService({ fallbackLang: 'en' }),
			IsMobileService,
			ViewerInspectorPanelService,
		],
	});
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	const fixture = TestBed.createComponent(HostComponent);
	fixture.detectChanges();
	return { fixture, inspector: TestBed.inject(ViewerInspectorPanelService) };
}

const sections = (fixture: ComponentFixture<HostComponent>) => [
	...(fixture.nativeElement as HTMLElement).querySelectorAll<SectionHost>(
		'[data-pptx-format-section]',
	),
];
const section = (fixture: ComponentFixture<HostComponent>, id: string) =>
	(fixture.nativeElement as HTMLElement).querySelector<SectionHost>(
		`[data-pptx-format-section="${id}"]`,
	)!;

describe('Angular Format pane sections', () => {
	it('orders a shape as Fill, Outline open, Effects closed, with bodies projected', () => {
		const { fixture } = open();
		const ids = sections(fixture).map((node) => node.dataset.pptxFormatSection);
		expect(ids.slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(section(fixture, 'fill').open).toBe(true);
		expect(section(fixture, 'fill').label).toBe('Fill');
		expect(section(fixture, 'outline').open).toBe(true);
		expect(section(fixture, 'effects').open).toBe(false);
		expect(section(fixture, 'sizePosition').open).toBe(false);
		expect(section(fixture, 'fill').querySelector('.fill-body')).not.toBeNull();
		expect(section(fixture, 'sizePosition').querySelector('.size-body')).not.toBeNull();
	});

	it('remembers a toggled section for the next element of the same category', () => {
		const { fixture } = open();
		section(fixture, 'fill').shadowRoot!.querySelector('summary')!.click();
		fixture.detectChanges();
		expect(section(fixture, 'fill').open).toBe(false);
		fixture.componentInstance.elements.set([shape('b')]);
		fixture.detectChanges();
		expect(section(fixture, 'fill').open).toBe(false);
	});

	it('opens the format view, expands and focuses a requested section', () => {
		vi.useFakeTimers({ toFake: ['requestAnimationFrame'] });
		try {
			const { fixture, inspector } = open();
			inspector.formatPanelClosed.set(true);
			inspector.formatPane.requestFormatPane('sizePosition', { focus: true });
			expect(inspector.formatPanelClosed()).toBe(false);
			fixture.detectChanges();
			expect(section(fixture, 'sizePosition').open).toBe(true);
			vi.advanceTimersToNextFrame();
			const target = section(fixture, 'sizePosition');
			expect(target.shadowRoot!.activeElement).toBe(target.shadowRoot!.querySelector('summary'));
		} finally {
			vi.useRealTimers();
		}
	});

	it('drops Size and position and tags differing values for a multi-selection', () => {
		const { fixture } = open();
		fixture.componentInstance.elements.set([shape('a', '#ff0000'), shape('b', '#00ff00')]);
		fixture.detectChanges();
		const ids = sections(fixture).map((node) => node.dataset.pptxFormatSection);
		expect(ids).not.toContain('sizePosition');
		expect(section(fixture, 'fill').mixed).toBe(true);
		expect(section(fixture, 'outline').mixed).toBe(false);
	});
});

describe('Angular inspector wiring', () => {
	const inspector = componentSource(import.meta.dirname, 'inspector-panel.component.ts');
	const slide = componentSource(import.meta.dirname, 'slide-default-inspector.component.ts');
	const menu = componentSource(import.meta.dirname, 'editor-context-menu.component.ts');
	const ribbon = componentSource(import.meta.dirname, 'ribbon.component.ts');

	it('puts the element panels in their sections and keeps the animation panel outside', () => {
		for (const id of [
			'font',
			'paragraph',
			'picture',
			'table',
			'chart',
			'smartArt',
			'media',
			'object',
			'fill',
			'outline',
			'effects',
			'arrange',
			'sizePosition',
			'altText',
			'action',
		]) {
			expect(inspector).toContain(`pptxFormatSection="${id}"`);
		}
		expect(inspector.indexOf('<pptx-animation-author-panel')).toBeGreaterThan(
			inspector.indexOf('</pptx-format-sections>'),
		);
		expect(slide).toContain('pptxFormatSection="slide"');
	});

	it('draws the Fill section from the shared fill section, the solid colour projected into it', () => {
		// The fill kind decides what shows: the solid colour only for a solid fill, the stops,
		// swatches and picture source only for theirs (`resolveShapeFillView`).
		const start = inspector.indexOf('<ng-template pptxFormatSection="fill">');
		const fill = inspector.slice(start, inspector.indexOf('</ng-template>', start));
		const section = fill.slice(
			fill.indexOf('<pptx-shape-fill-section'),
			fill.indexOf('</pptx-shape-fill-section>'),
		);
		expect(section).toContain('id="insp-fill"');
		expect(section).toContain('<pptx-theme-color-swatch-grid');
		expect(fill).not.toContain('<pptx-gradient-picker');
		expect(fill).not.toContain('<pptx-pattern-fill-panel');
	});

	it('routes the context menu and the Home launchers through requestFormatPane', () => {
		expect(menu).toContain('formatPane.requestFormatPane(');
		expect(menu).toContain('formatSectionForAnchor(anchor)');
		expect(ribbon).toContain('HOME_LAUNCHER_SECTIONS[group]');
	});
});
