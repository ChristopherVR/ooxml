/**
 * The shared element render error contract (`shared/test-utils/element-render-error-contract.mjs`)
 * run against the Angular canvas and slides pane blocks over one `ViewerBuildingBlocks`.
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
import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../shared/test-utils/building-blocks-contract.mjs';
import { runElementRenderErrorContract } from '../../../shared/test-utils/element-render-error-contract.mjs';
import { resolveViewerComponentResources } from '../viewer/component-resources.test-support';
import { declareSignalFields } from '../viewer/signal-inputs.test-support';
import { BuildingBlocksCanvasComponent } from './canvas.component';
import { BuildingBlocksSlidesPaneComponent } from './slides-pane.component';
import type { ViewerBuildingBlocksOptions } from './types';
import {
	createViewerBuildingBlocks,
	VIEWER_BUILDING_BLOCKS_PROVIDERS,
} from './viewer-building-blocks';

@Component({
	selector: 'test-element-errors-host',
	standalone: true,
	providers: [...VIEWER_BUILDING_BLOCKS_PROVIDERS],
	imports: [BuildingBlocksSlidesPaneComponent, BuildingBlocksCanvasComponent],
	template: `
		<pptx-slides-pane [blocks]="blocks" style="display: block; width: 200px" />
		<pptx-building-blocks-canvas
			[blocks]="blocks"
			style="display: block; width: 640px; height: 360px"
		/>
	`,
})
class ElementErrorsHostComponent {
	readonly options = signal<ViewerBuildingBlocksOptions>({});
	readonly blocks = createViewerBuildingBlocks(() => this.options());
}

const tick = () =>
	new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});

beforeAll(async () => {
	declareSignalFields(ElementErrorsHostComponent, path.resolve(import.meta.dirname, '..'));
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

let fixture: ComponentFixture<ElementErrorsHostComponent> | undefined;

afterEach(() => {
	fixture = undefined;
	TestBed.resetTestingModule();
	document.body.replaceChildren();
});

const settle = async () => {
	fixture?.detectChanges();
	await tick();
	fixture?.detectChanges();
};

function angularHarness(): BuildingBlocksContractHarness {
	return {
		name: 'angular',
		toolbar: false,
		async mount(target, options: BuildingBlocksContractOptions) {
			TestBed.resetTestingModule();
			TestBed.configureTestingModule({
				providers: [provideTranslateService({ fallbackLang: 'en' })],
			});
			TestBed.inject(TranslateService).setTranslation('en', translationsEn);
			TestBed.inject(TranslateService).use('en');
			fixture = TestBed.createComponent(ElementErrorsHostComponent);
			target.append(fixture.nativeElement);
			fixture.componentInstance.options.set(options as ViewerBuildingBlocksOptions);
			fixture.detectChanges();
			await settle();
		},
		async update(options) {
			fixture!.componentInstance.options.set(options as ViewerBuildingBlocksOptions);
			await settle();
		},
		actions: () => fixture!.componentInstance.blocks.actions(),
		state: () => fixture!.componentInstance.blocks.state(),
		clickPictures: settle,
		settle,
		destroy: () => {
			fixture?.destroy();
			fixture = undefined;
		},
	};
}

describe('angular building blocks element render errors', () => {
	it('passes the shared element render error contract', async () => {
		await runElementRenderErrorContract({
			harness: angularHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 90000);
});
