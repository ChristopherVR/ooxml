/**
 * ribbon-tab-list.component.ts: the ribbon's tab strip row, split out of
 * {@link RibbonComponent} (which was well over this repo's 300-LOC file cap).
 *
 * Renders the shared `pptx-ui-ribbon-tabs` (the tab row Word, Excel and Visio draw), fed from
 * `buildRibbonTabRow`, with the tab-row actions in its actions slot: Record, then Comments and
 * Share drawn by the shared `pptx-ui-ribbon-actions`, then the Ribbon Display Options menu button.
 */
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
	output,
} from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { LucideChevronDown } from '@lucide/angular';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import {
	buildRibbonTabRow,
	buildTabRowActionsState,
	RIBBON_DISPLAY_OPTIONS_LABEL_KEY,
	TAB_ROW_ACTION_CLASSES,
} from 'ooxml-ui/pptx';
import type {
	RibbonAddInTab,
	RibbonContextualTabId,
	RibbonDisplayController,
	RibbonDisplayState,
	RibbonTabSelectDetail,
	ToolbarActionId,
} from 'ooxml-ui/pptx';
import { map, merge, startWith } from 'rxjs';

import { RibbonDisplayMenuComponent } from './ribbon-display-menu.component';
import type { RibbonTab } from './ribbon-types';
import { toolbarVisibility } from './toolbar-visibility';
import { ViewerOptionsService } from './viewer-options.service';

@Component({
	selector: 'pptx-ribbon-tab-list',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe, LucideChevronDown, RibbonDisplayMenuComponent],
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		<pptx-ui-ribbon-tabs
			data-pptx-chrome="ribbon-tabs"
			[state]="tabRow()"
			(tab-select)="onTabSelect($event)"
			(collapse-toggle)="displayController()?.toggleCollapsed()"
		>
			<!-- Tab-row right actions (Record, Comments, Share), mirroring React's TabRowActions -->
			<div slot="actions" class="flex shrink-0 items-center gap-1 pr-1">
				@if (canEdit() && !toolbar.isHidden('record')) {
					<button
						type="button"
						[class]="tra.record"
						[title]="'pptx.titleBar.record' | translate"
						[attr.aria-label]="'pptx.titleBar.record' | translate"
						(click)="record.emit()"
					>
						<span [class]="tra.recordDot" aria-hidden="true"></span>
						<span>{{ 'pptx.titleBar.record' | translate }}</span>
					</button>
				}
				<pptx-ui-ribbon-actions
					data-pptx-chrome="tab-row-actions"
					[state]="actionsState()"
					(comments-toggle)="toggleComments.emit()"
					(share-request)="share.emit()"
				></pptx-ui-ribbon-actions>

				<button
					type="button"
					data-pptx-chrome="ribbon-display-options"
					class="mr-1 inline-flex h-[22px] w-[22px] shrink-0 items-center justify-center rounded text-muted-foreground transition-colors hover:bg-accent/30 hover:text-foreground"
					aria-haspopup="menu"
					[attr.aria-expanded]="menuOpen()"
					[attr.aria-label]="displayOptionsKey | translate"
					[title]="displayOptionsKey | translate"
					(click)="displayController()?.toggleMenu()"
				>
					<svg
						lucideChevronDown
						class="h-3 w-3 transition-transform"
						[class.rotate-180]="effectiveMode() === 'tabs'"
						aria-hidden="true"
					></svg>
				</button>
			</div>
		</pptx-ui-ribbon-tabs>
		<!-- After the tab row, so the menu hangs under the tabs rather than inside the scroller. -->
		<pptx-ribbon-display-menu [state]="displayState()" [controller]="displayController()" />
	`,
})
export class RibbonTabListComponent {
	readonly activeTab = input.required<RibbonTab>();
	readonly canEdit = input<boolean>(false);
	readonly collabConnected = input<boolean>(false);
	readonly connectedCount = input<number>(0);
	/** Comments pane state and the current slide's comment count (Comments' badge). */
	readonly commentsOpen = input<boolean>(false);
	readonly commentCount = input<number>(0);
	/** Ribbon Display Options state (null: the docked ribbon, menu closed). */
	readonly displayState = input<RibbonDisplayState | null>(null);
	/** The shared Ribbon Display Options controller the menu and the tabs call. */
	readonly displayController = input<RibbonDisplayController | null>(null);
	/** Toolbar tabs/buttons the host wants hidden (filters the tab strip; gates Record/Share). */
	readonly hiddenActions = input<ToolbarActionId[]>([]);
	/** Contextual tabs the selection brings up, appended after the fixed tabs. */
	readonly contextualTabs = input<readonly RibbonContextualTabId[]>([]);
	/** The host's tabs (`ribbonAddIns`), shown after the fixed tabs. */
	readonly addInTabs = input<readonly RibbonAddInTab[]>([]);
	/** The host tab being shown; no fixed or contextual tab is selected while one is. */
	readonly activeAddIn = input<string | null>(null);

	readonly selectTab = output<RibbonTab>();
	readonly selectAddIn = output<string>();
	readonly record = output<void>();
	readonly share = output<void>();
	readonly toggleComments = output<void>();

	protected readonly toolbar = toolbarVisibility(this.hiddenActions);
	protected readonly displayOptionsKey = RIBBON_DISPLAY_OPTIONS_LABEL_KEY;
	protected readonly menuOpen = computed(() => this.displayState()?.menuOpen ?? false);
	protected readonly effectiveMode = computed(() => this.displayState()?.effectiveMode ?? 'always');
	protected readonly tra = TAB_ROW_ACTION_CLASSES;

	private readonly translate = inject(TranslateService);
	/** Changes on language/dictionary updates so the element's labels re-translate. */
	private readonly translations = toSignal(
		merge(this.translate.onLangChange, this.translate.onTranslationChange).pipe(
			map(() => Date.now()),
			startWith(0),
		),
		{ initialValue: 0 },
	);
	/** Comments and Share for the shared element, from the shared descriptor. */
	protected readonly actionsState = computed(() => {
		this.translations();
		return buildTabRowActionsState({
			translate: (key, params) => this.translate.instant(key, params) as string,
			showComments: true,
			commentsOpen: this.commentsOpen(),
			commentCount: this.commentCount(),
			showShare: !this.toolbar.isHidden('share'),
			isCollaborating: this.collabConnected(),
			collaboratorCount: this.connectedCount(),
		});
	});
	/** Optional so the tab strip renders outside a full viewer host too. */
	private readonly viewerOpts = inject(ViewerOptionsService, { optional: true });

	/** The tab row as the shared element draws it. */
	protected readonly tabRow = computed(() => {
		this.translations();
		return buildRibbonTabRow({
			translate: (key) => this.translate.instant(key) as string,
			active: this.activeAddIn() ?? this.activeTab(),
			hiddenActions: this.hiddenActions(),
			addInTabs: this.addInTabs(),
			contextualTabs: this.contextualTabs(),
			screenTip: (label) => (this.viewerOpts ? this.viewerOpts.screenTip(label) : label),
		});
	});

	/** A tab was chosen: a host tab goes to `selectAddIn`, the others to `selectTab`. */
	protected onTabSelect(event: Event): void {
		const { tab, kind } = (event as CustomEvent<RibbonTabSelectDetail>).detail;
		if (kind === 'addIn') {
			this.selectAddIn.emit(tab);
		} else {
			this.selectTab.emit(tab as RibbonTab);
		}
	}
}
