import {
	ChangeDetectionStrategy,
	Component,
	effect,
	ElementRef,
	inject,
	signal,
	viewChild,
} from '@angular/core';
import type { AfterViewInit } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	BuildingBlocksCanvasComponent,
	BuildingBlocksSlidesPaneComponent,
	BuildingBlocksToolbarComponent,
	createViewerBuildingBlocks,
	translationsEn,
	VIEWER_BUILDING_BLOCKS_PROVIDERS,
} from 'pptx-angular-viewer';
import type { ViewerBuildingBlocksOptions } from 'pptx-angular-viewer';

import { mountBuildingBlocksPanel } from '../../shared/building-blocks-panel';

/**
 * `?blocks=1`: a custom shell composed from the Angular building blocks (toolbar, slides pane,
 * canvas) with the shared host panel that finishes their workflows. The framework-neutral
 * `e2e/pptx/building-blocks.spec.ts` drives it.
 */
@Component({
	selector: 'app-root',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	providers: [...VIEWER_BUILDING_BLOCKS_PROVIDERS],
	imports: [
		BuildingBlocksToolbarComponent,
		BuildingBlocksSlidesPaneComponent,
		BuildingBlocksCanvasComponent,
	],
	template: `
		<section #panel></section>
		<pptx-building-blocks-toolbar [blocks]="blocks" />
		<div style="display: flex; height: calc(100vh - 160px)">
			<pptx-slides-pane [blocks]="blocks" style="width: 200px; overflow: auto" />
			<pptx-building-blocks-canvas [blocks]="blocks" style="flex: 1; position: relative" />
		</div>
	`,
})
export class BuildingBlocksDemoComponent implements AfterViewInit {
	private readonly options = signal<ViewerBuildingBlocksOptions>({ canEdit: false });
	readonly blocks = createViewerBuildingBlocks(() => this.options());
	private readonly panel = viewChild.required<ElementRef<HTMLElement>>('panel');
	private host: { refresh(): void } | undefined;

	constructor() {
		const translate = inject(TranslateService);
		translate.setTranslation('en', translationsEn);
		translate.use('en');
		void fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`)
			.then((response) => response.arrayBuffer())
			.then((bytes) => this.options.update((o) => ({ ...o, content: new Uint8Array(bytes) })));
		// Refresh the host panel after every change of the blocks' state or actions.
		effect(() => {
			this.blocks.actions();
			this.blocks.state();
			this.host?.refresh();
		});
	}

	ngAfterViewInit(): void {
		this.host = mountBuildingBlocksPanel(this.panel().nativeElement, {
			actions: () => this.blocks.actions(),
			state: () => this.blocks.state(),
			setCanEdit: (canEdit) => this.options.update((o) => ({ ...o, canEdit })),
			setCollaboration: (collaboration) =>
				this.options.update((o) => ({
					...o,
					collaboration: collaboration as ViewerBuildingBlocksOptions['collaboration'],
				})),
		});
		this.host.refresh();
	}
}
