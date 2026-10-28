/**
 * The slide thumbnail rail (`SlidesPanelComponent`) as a host-driven building
 * block: real thumbnails (real `SlideCanvasComponent`s) rendered from host
 * inputs, the shared field and table style contexts, navigation without an
 * edit, the Layout hand-off, multi-selection commands and read-only mode.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { Component, Injector, signal } from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import { By } from '@angular/platform-browser';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { SlidesPaneLayoutGalleryRequest } from 'ooxml-ui/pptx';
import type { PptxSection, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import * as publicApi from '../public-api';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { EditorStateService } from './editor-state.service';
import { FieldContextService } from './field-context.service';
import { LoadContentService } from './load-content.service';
import { declareSignalFields } from './signal-inputs.test-support';
import { SlideCanvasComponent } from './slide-canvas.component';
import { createSlidesPanelEditorCommands } from './slides-panel-commands';
import { SlidesPanelComponent } from './slides-panel.component';
import type { TemplateElementsBySlideId } from './template-mode';

const FIXTURE = path.resolve(
	import.meta.dirname,
	'../../../../../../e2e/pptx/fixtures/first-slide-number.pptx',
);

/** A host that owns the deck state and places the rail in its own shell. */
@Component({
	selector: 'test-rail-host',
	standalone: true,
	imports: [SlidesPanelComponent],
	template: `
		<pptx-slides-panel
			[slides]="slides()"
			[templateElementsBySlideId]="templates()"
			[sections]="sections()"
			[canvasSize]="{ width: 960, height: 540 }"
			[activeIndex]="active()"
			[canEdit]="canEdit()"
			[firstSlideNumber]="firstSlideNumber()"
			(select)="onSelect($event)"
			(addSlide)="events.push(['addSlide', $event])"
			(duplicateSlides)="events.push(['duplicateSlides', $event])"
			(deleteSlides)="events.push(['deleteSlides', $event])"
			(openLayoutGallery)="onLayout($event)"
			(openLayoutForSlide)="events.push(['openLayoutForSlide', $event])"
			(addSectionAt)="events.push(['addSectionAt', $event])"
			(toggleSectionCollapse)="events.push(['toggleSectionCollapse', $event])"
		/>
	`,
})
class RailHostComponent {
	readonly slides = signal<readonly PptxSlide[]>([]);
	readonly templates = signal<TemplateElementsBySlideId>({});
	readonly sections = signal<readonly PptxSection[]>([]);
	readonly active = signal(0);
	readonly canEdit = signal(true);
	readonly firstSlideNumber = signal<number | undefined>(undefined);
	readonly events: [string, unknown][] = [];
	onSelect(index: number): void {
		this.events.push(['select', index]);
		this.active.set(index);
	}
	onLayout(request: SlidesPaneLayoutGalleryRequest): void {
		this.events.push(['openLayoutGallery', request]);
	}
}

/** The rail editor-bound, as `PowerPointViewerComponent` places it. */
@Component({
	selector: 'test-viewer-rail',
	standalone: true,
	imports: [SlidesPanelComponent, SlideCanvasComponent],
	template: ` <pptx-slides-panel [canvasSize]="{ width: 960, height: 540 }" [activeIndex]="0" /> `,
})
class ViewerRailComponent {}

beforeAll(async () => {
	declareSignalFields(RailHostComponent, import.meta.dirname);
	declareSignalFields(ViewerRailComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

/** Load the fixture outside any component, as a host's own state would. */
async function loadFixture(): Promise<LoadContentService> {
	const injector = Injector.create({
		providers: [{ provide: LoadContentService, useClass: LoadContentService }],
	});
	const loader = injector.get(LoadContentService);
	await loader.load(new Uint8Array(readFileSync(FIXTURE)));
	return loader;
}

function configure(providers: unknown[] = []): void {
	TestBed.configureTestingModule({
		providers: [provideTranslateService({ fallbackLang: 'en' }), ...(providers as never[])],
	});
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
}

function blankDeck(count: number): PptxSlide[] {
	return Array.from(
		{ length: count },
		(_, i) => ({ id: `s${i}`, rId: `s${i}`, slideNumber: i + 1, elements: [] }) as PptxSlide,
	);
}

function rows(fixture: ComponentFixture<unknown>): HTMLButtonElement[] {
	return [
		...(fixture.nativeElement as HTMLElement).querySelectorAll<HTMLButtonElement>(
			'[data-pptx-chrome="slide-row"]',
		),
	];
}

function frameTexts(fixture: ComponentFixture<unknown>): string[] {
	return [
		...(fixture.nativeElement as HTMLElement).querySelectorAll('[data-pptx-chrome="slide-frame"]'),
	].map((frame) => frame.textContent?.replace(/\s+/gu, ' ').trim() ?? '');
}

/** Right-click card `index`, then pick `command` from the shared menu element. */
function runMenuCommand(fixture: ComponentFixture<unknown>, index: number, command: string): void {
	const card = (fixture.nativeElement as HTMLElement).querySelectorAll(
		'[data-pptx-chrome="slide-wrapper"]',
	)[index]!;
	card.dispatchEvent(
		new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 120 }),
	);
	fixture.detectChanges();
	const menu = (fixture.nativeElement as HTMLElement).querySelector('pptx-ui-context-menu');
	expect(menu).not.toBeNull();
	menu!.dispatchEvent(new CustomEvent('menu-request', { detail: { id: command } }));
	fixture.detectChanges();
}

function mountHost(slides: readonly PptxSlide[], providers: unknown[] = []) {
	configure(providers);
	const fixture = TestBed.createComponent(RailHostComponent);
	fixture.componentInstance.slides.set(slides);
	fixture.detectChanges();
	return fixture;
}

describe('slides panel building block', () => {
	it('is exported from the package root', () => {
		expect(publicApi.SlidesPanelComponent).toBe(SlidesPanelComponent);
		expect(publicApi.createSlidesPanelEditorCommands).toBe(createSlidesPanelEditorCommands);
	});

	it('numbers host-driven thumbnails from firstSlideNum with no viewer state in scope', async () => {
		const loader = await loadFixture();
		const fixture = mountHost(loader.slides());
		fixture.componentInstance.firstSlideNumber.set(loader.firstSlideNumber());
		fixture.detectChanges();
		expect(loader.firstSlideNumber()).toBe(45);
		const texts = frameTexts(fixture);
		expect(texts).toHaveLength(3);
		expect(texts[0]).toContain('Slide 45');
		expect(texts[1]).toContain('Slide 46');
		expect(texts[2]).toContain('Slide 47');
	}, 30_000);

	it('gives every thumbnail the canvas field context, re-pointed at its own slide', async () => {
		configure([LoadContentService, FieldContextService, EditorStateService]);
		const loader = TestBed.inject(LoadContentService);
		await loader.load(new Uint8Array(readFileSync(FIXTURE)));
		const editor = TestBed.inject(EditorStateService);
		editor.setSlides(loader.slides(), loader.sections());
		const fixture = TestBed.createComponent(ViewerRailComponent);
		fixture.detectChanges();

		const canvasFields = TestBed.inject(FieldContextService);
		const panel = fixture.debugElement.query(By.directive(SlidesPanelComponent))
			.componentInstance as SlidesPanelComponent;
		expect(panel.thumbnailContexts.fields.deckContext()).toStrictEqual(canvasFields.deckContext());
		const thumbnails = fixture.debugElement
			.queryAll(By.directive(SlideCanvasComponent))
			.map((debug) => debug.componentInstance as SlideCanvasComponent);
		expect(thumbnails).toHaveLength(3);
		thumbnails.forEach((thumbnail, i) => {
			expect(thumbnail.fieldContext()).toStrictEqual(canvasFields.forSlide(editor.slides()[i]));
		});
		expect(frameTexts(fixture).map((text) => /Slide \d+/u.exec(text)?.[0])).toStrictEqual([
			'Slide 45',
			'Slide 46',
			'Slide 47',
		]);
	}, 30_000);

	it('emits the DOM markers the cross-binding rail specs use', () => {
		const fixture = mountHost(blankDeck(3));
		const all = rows(fixture);
		expect(all).toHaveLength(3);
		expect(all[0]!.getAttribute('aria-current')).toBe('true');
		expect(all[0]!.getAttribute('aria-label')).toBe('Go to slide 1');
		expect(all[1]!.hasAttribute('aria-current')).toBeFalsy();
		expect(all[0]!.querySelector('[data-pptx-chrome="slide-frame"]')).not.toBeNull();
		const footer = (fixture.nativeElement as HTMLElement).querySelector(
			'[data-pptx-chrome="slide-footer"]',
		);
		expect(footer?.querySelectorAll('button')).toHaveLength(1);
	});

	it('moves the active slide on click without editing the deck', () => {
		configure([EditorStateService]);
		const editor = TestBed.inject(EditorStateService);
		editor.setSlides(blankDeck(3));
		const deck = editor.slides();
		const fixture = TestBed.createComponent(RailHostComponent);
		fixture.componentInstance.slides.set(editor.slides());
		fixture.detectChanges();

		rows(fixture)[2]!.click();
		fixture.detectChanges();
		expect(fixture.componentInstance.events).toStrictEqual([['select', 2]]);
		expect(rows(fixture)[2]!.getAttribute('aria-current')).toBe('true');
		expect(editor.slides()).toBe(deck);
		expect(editor.dirty()).toBeFalsy();
		expect(editor.canUndo()).toBeFalsy();
	});

	it('makes the slide active, then hands the Layout entry to the host', () => {
		const fixture = mountHost(blankDeck(3));
		runMenuCommand(fixture, 1, 'layout');
		expect(fixture.componentInstance.events).toStrictEqual([
			['select', 1],
			['openLayoutGallery', { slideIndex: 1, x: 40, y: 120 }],
		]);
		expect(rows(fixture)[1]!.getAttribute('aria-current')).toBe('true');
	});

	it('duplicates and deletes every selected slide through the editor commands', () => {
		configure([EditorStateService]);
		const editor = TestBed.inject(EditorStateService);
		editor.setSlides(blankDeck(4));
		const commands = createSlidesPanelEditorCommands(editor);
		const fixture = TestBed.createComponent(RailHostComponent);
		const host = fixture.componentInstance;
		host.slides.set(editor.slides());
		fixture.detectChanges();

		rows(fixture)[1]!.click();
		rows(fixture)[3]!.dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
		runMenuCommand(fixture, 3, 'duplicate');
		const duplicate = host.events.find(([name]) => name === 'duplicateSlides');
		expect(duplicate?.[1]).toStrictEqual([1, 3]);
		commands.duplicateSlides(duplicate![1] as number[]);
		const ids = editor.slides().map((slide) => slide.id);
		expect(ids).toHaveLength(6);
		expect([ids[0], ids[1], ids[3], ids[4]]).toStrictEqual(['s0', 's1', 's2', 's3']);
		expect(editor.slides()[2]!.elements).toStrictEqual(editor.slides()[1]!.elements);
		expect(ids[5]).not.toBe('s3');

		host.slides.set(editor.slides());
		host.events.length = 0;
		fixture.detectChanges();
		runMenuCommand(fixture, 4, 'delete');
		const remove = host.events.find(([name]) => name === 'deleteSlides');
		expect(remove?.[1]).toStrictEqual([1, 4]);
		commands.deleteSlides(remove![1] as number[]);
		expect(editor.slides().map((slide) => slide.id)).toStrictEqual(['s0', ids[2], 's2', ids[5]]);
	});

	it('emits the footer add for the host and leaves the deck alone', () => {
		configure([EditorStateService]);
		const editor = TestBed.inject(EditorStateService);
		editor.setSlides(blankDeck(2));
		const fixture = TestBed.createComponent(RailHostComponent);
		fixture.componentInstance.slides.set(editor.slides());
		fixture.componentInstance.active.set(1);
		fixture.detectChanges();
		(fixture.nativeElement as HTMLElement)
			.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button')!
			.click();
		expect(fixture.componentInstance.events).toStrictEqual([['addSlide', 1]]);
		expect(editor.slides()).toHaveLength(2);
		createSlidesPanelEditorCommands(editor).addSlideAfter(1);
		expect(editor.slides()).toHaveLength(3);
	});

	it('still edits the enclosing editor itself when editor-bound', () => {
		configure([EditorStateService]);
		const editor = TestBed.inject(EditorStateService);
		editor.setSlides(blankDeck(2));
		const fixture = TestBed.createComponent(ViewerRailComponent);
		fixture.detectChanges();
		(fixture.nativeElement as HTMLElement)
			.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button')!
			.click();
		expect(editor.slides()).toHaveLength(3);
		expect(editor.dirty()).toBeTruthy();
	});

	it('groups under section headers and emits section commands', () => {
		const sectioned = blankDeck(3).map((slide, i) => ({ ...slide, sectionId: i < 2 ? 'a' : 'b' }));
		const fixture = mountHost(sectioned);
		fixture.componentInstance.sections.set([
			{ id: 'a', name: 'Intro', slideIds: ['s0', 's1'] },
			{ id: 'b', name: 'Close', slideIds: ['s2'] },
		] as PptxSection[]);
		fixture.detectChanges();
		const headers = (fixture.nativeElement as HTMLElement).querySelectorAll(
			'[data-pptx-chrome="section-header"]',
		);
		expect([...headers].map((header) => header.textContent?.trim())).toStrictEqual([
			'▾Intro',
			'▾Close',
		]);
		headers[1]!.querySelector('button')!.click();
		expect(fixture.componentInstance.events).toStrictEqual([['toggleSectionCollapse', 'b']]);
	});

	it('shows no footer and opens no menu when read-only', () => {
		const fixture = mountHost(blankDeck(3));
		fixture.componentInstance.canEdit.set(false);
		fixture.detectChanges();
		const root = fixture.nativeElement as HTMLElement;
		expect(root.querySelector('[data-pptx-chrome="slide-footer"]')).toBeNull();
		root
			.querySelector('[data-pptx-chrome="slide-wrapper"]')!
			.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
		fixture.detectChanges();
		expect(root.querySelector('pptx-ui-context-menu')).toBeNull();
		expect(
			root.querySelector('[data-pptx-chrome="slide-wrapper"]')!.getAttribute('draggable'),
		).toBe('false');
		const onKey = vi.fn();
		const panel = fixture.debugElement.query(By.directive(SlidesPanelComponent))
			.componentInstance as SlidesPanelComponent;
		panel.addSlideAfter.subscribe(onKey);
		root
			.querySelector('pptx-slides-panel')!
			.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(onKey).not.toHaveBeenCalled();
		expect(rows(fixture)).toHaveLength(3);
	});
});
