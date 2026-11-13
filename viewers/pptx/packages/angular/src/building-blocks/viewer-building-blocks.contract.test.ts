/**
 * The shared building-blocks parity contract (`shared/test-utils/building-blocks-contract.mjs`)
 * run against the Angular blocks: the real toolbar, slides pane and canvas components over one
 * `ViewerBuildingBlocks`, plus the behaviour that is specific to this binding.
 */
import path from 'node:path';

import { Component, signal } from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import {
	buildingBlocksContractDeck,
	runBuildingBlocksContract,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import * as publicApi from '../public-api';
import { resolveViewerComponentResources } from '../viewer/component-resources.test-support';
import { declareSignalFields } from '../viewer/signal-inputs.test-support';
import { BuildingBlocksCanvasComponent } from './canvas.component';
import { FIND_MATCH_ATTRIBUTE } from 'ooxml-ui/pptx';
import { BuildingBlocksSlidesPaneComponent } from './slides-pane.component';
import { BuildingBlocksToolbarComponent } from './toolbar.component';
import type { ViewerBuildingBlocksOptions } from './types';
import {
	createViewerBuildingBlocks,
	VIEWER_BUILDING_BLOCKS_PROVIDERS,
} from './viewer-building-blocks';

@Component({
	selector: 'test-blocks-host',
	standalone: true,
	providers: [...VIEWER_BUILDING_BLOCKS_PROVIDERS],
	imports: [
		BuildingBlocksToolbarComponent,
		BuildingBlocksSlidesPaneComponent,
		BuildingBlocksCanvasComponent,
	],
	template: `
		<pptx-building-blocks-toolbar [blocks]="blocks" />
		<pptx-slides-pane [blocks]="blocks" style="display: block; width: 200px" />
		<pptx-building-blocks-canvas
			[blocks]="blocks"
			style="display: block; width: 640px; height: 360px"
		/>
	`,
})
class BlocksHostComponent {
	readonly options = signal<ViewerBuildingBlocksOptions>({});
	readonly blocks = createViewerBuildingBlocks(() => this.options());
}

const tick = () =>
	new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});

beforeAll(async () => {
	declareSignalFields(BlocksHostComponent, path.resolve(import.meta.dirname, '..'));
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

let fixture: ComponentFixture<BlocksHostComponent> | undefined;

afterEach(() => {
	fixture = undefined;
	TestBed.resetTestingModule();
	document.body.replaceChildren();
});

function create(options: ViewerBuildingBlocksOptions, target?: HTMLElement) {
	TestBed.configureTestingModule({
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
	const created = TestBed.createComponent(BlocksHostComponent);
	fixture = created;
	(target ?? document.body).append(created.nativeElement);
	created.componentInstance.options.set(options);
	created.detectChanges();
	return created;
}

const settle = async () => {
	fixture?.detectChanges();
	await tick();
	fixture?.detectChanges();
};

function angularHarness(): BuildingBlocksContractHarness {
	let host: HTMLElement | undefined;
	return {
		name: 'angular',
		toolbar: true,
		async mount(target, options: BuildingBlocksContractOptions) {
			host = target;
			create(options as ViewerBuildingBlocksOptions, target);
			await settle();
		},
		async update(options) {
			fixture!.componentInstance.options.set(options as ViewerBuildingBlocksOptions);
			await settle();
		},
		actions: () => fixture!.componentInstance.blocks.actions(),
		state: () => fixture!.componentInstance.blocks.state(),
		async clickPictures() {
			host!.querySelector<HTMLButtonElement>('[data-control="insertImage"]')!.click();
			await settle();
		},
		settle,
		destroy: () => fixture?.destroy(),
	};
}

async function mounted(extra: Partial<ViewerBuildingBlocksOptions> = {}) {
	const content = await buildingBlocksContractDeck(PptxHandler);
	const created = create({ content, canEdit: true, ...extra });
	await vi.waitFor(async () => {
		await settle();
		expect(created.componentInstance.blocks.state().slideCount).toBe(2);
	});
	return { created, blocks: created.componentInstance.blocks, content };
}

describe('angular building blocks', () => {
	it('passes the shared building-blocks contract', async () => {
		await runBuildingBlocksContract({
			harness: angularHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 60000);

	it('is exported from the package root', () => {
		expect(publicApi.ViewerBuildingBlocks).toBeDefined();
		expect(publicApi.createViewerBuildingBlocks).toBe(createViewerBuildingBlocks);
		expect(publicApi.BuildingBlocksToolbarComponent).toBe(BuildingBlocksToolbarComponent);
		expect(publicApi.BuildingBlocksSlidesPaneComponent).toBe(BuildingBlocksSlidesPaneComponent);
		expect(publicApi.BuildingBlocksCanvasComponent).toBe(BuildingBlocksCanvasComponent);
	});

	it('mounts no ribbon or dialogs and reports not-ready before the deck loads', async () => {
		const content = await buildingBlocksContractDeck(PptxHandler);
		const created = create({ content, canEdit: true });
		expect(created.componentInstance.blocks.actions().availability.insertImage.available).toBe(
			false,
		);
		expect(created.componentInstance.blocks.state()).toMatchObject({
			mode: 'edit',
			loading: true,
			slideCount: 0,
		});
		await vi.waitFor(async () => {
			await settle();
			expect(created.componentInstance.blocks.state().slideCount).toBe(2);
		});
		const root = created.nativeElement as HTMLElement;
		expect(root.querySelector('[data-pptx-viewport]')).not.toBeNull();
		expect(root.querySelector('pptx-ribbon')).toBeNull();
	}, 30000);

	it('outlines the elements holding the host find bar matches', async () => {
		const { created, blocks } = await mounted();
		blocks.actions().findReplace.setOpen(true);
		blocks.actions().findReplace.setQuery('alpha');
		await vi.waitFor(async () => {
			await settle();
			expect(
				(created.nativeElement as HTMLElement).querySelector(`[${FIND_MATCH_ATTRIBUTE}="current"]`),
			).not.toBeNull();
		});
		blocks.actions().findReplace.setOpen(false);
		await vi.waitFor(async () => {
			await settle();
			expect(
				(created.nativeElement as HTMLElement).querySelector(`[${FIND_MATCH_ATTRIBUTE}]`),
			).toBeNull();
		});
	}, 30000);

	it('reports unknown-layout when the handler rejects a listed layout', async () => {
		const { blocks } = await mounted();
		const layout = blocks.actions().layout.available[0]!;
		expect(layout).toBeDefined();
		vi.spyOn(blocks.shell.loader.getHandler()!, 'applyLayoutToSlide').mockRejectedValueOnce(
			new Error('broken layout'),
		);
		expect(await blocks.actions().layout.apply(layout.path)).toEqual({
			ok: false,
			code: 'unknown-layout',
		});
		vi.restoreAllMocks();
	}, 30000);

	it('applies new callbacks and permission without a remount', async () => {
		const first = vi.fn();
		const { created, blocks, content } = await mounted({ onDirtyChange: first });
		const second = vi.fn();
		created.componentInstance.options.set({ content, canEdit: true, onDirtyChange: second });
		await settle();
		expect(blocks.actions().smartArt.insert('basicBlockList', ['A'])).toEqual({ ok: true });
		await vi.waitFor(async () => {
			await settle();
			expect(second).toHaveBeenCalledWith(true);
		});
		expect(first).not.toHaveBeenCalled();
		created.componentInstance.options.set({ content, canEdit: false, onDirtyChange: second });
		await settle();
		expect(blocks.state().mode).toBe('preview');
	}, 30000);

	it('turns the find and shortcut chords into the host find bar and shortcut help', async () => {
		const onRequest = vi.fn();
		const { blocks } = await mounted({ onRequest });
		document.body.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true }),
		);
		await vi.waitFor(async () => {
			await settle();
			expect(blocks.actions().findReplace.open).toBe(true);
		});
		document.body.dispatchEvent(
			new KeyboardEvent('keydown', { key: '/', ctrlKey: true, bubbles: true, cancelable: true }),
		);
		await vi.waitFor(async () => {
			await settle();
			expect(blocks.actions().shortcuts.open).toBe(true);
		});
		expect(onRequest).toHaveBeenCalledWith({ kind: 'shortcuts', open: true });
	}, 30000);
});
