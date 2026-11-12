/**
 * ribbon-tab-list.component.ts: the ribbon's tab strip row, split out of
 * {@link RibbonComponent} (which was well over this repo's 300-LOC file cap).
 *
 * Renders: the scrollable tab strip (File/Home/Insert/.../Help, filtered by
 * `hiddenActions`), the pinned tab-row actions (Record, then Comments and Share
 * drawn by the shared `pptx-ui-ribbon-actions`, the element Word and Excel put
 * at the same place), and the Ribbon Display Options menu button. Behaviour and markup are unchanged
 * from the original inline tab-bar `<div>` in `ribbon.component.ts`.
 */
import { NgClass } from '@angular/common';
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
	buildTabRowActionsState,
	contextualTabLabelKey,
	filterVisibleTabs,
	RIBBON_DISPLAY_OPTIONS_LABEL_KEY,
	TAB_ROW_ACTION_CLASSES,
	TOOLBAR_TABS,
} from 'ooxml-ui/pptx';
import type {
	RibbonAddInTab,
	RibbonContextualTabId,
	RibbonDisplayController,
	RibbonDisplayState,
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
	imports: [NgClass, TranslatePipe, LucideChevronDown, RibbonDisplayMenuComponent],
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		<div
			role="tablist"
			data-pptx-chrome="ribbon-tabs"
			class="flex items-center border-b border-border/60 px-1"
		>
			<!-- Scrollable tab strip: on narrow widths the tabs scroll instead of
			     clipping (mirrors React's max-md:overflow-x-auto scrollbar-none),
			     while the Record/Share actions and collapse toggle stay pinned. -->
			<div
				class="flex min-w-0 flex-1 items-center overflow-x-auto pptx-scrollbar-none"
				data-pptx-chrome="ribbon-tab-scroll"
			>
				@for (t of visibleTabs(); track t.id) {
					<button
						type="button"
						role="tab"
						[attr.aria-selected]="!activeAddIn() && activeTab() === t.id"
						[title]="tabTip(t.labelKey)"
						(click)="selectTab.emit(t.id)"
						(dblclick)="tabDoubleClick(t.id)"
						class="relative whitespace-nowrap px-3.5 py-2 text-[12px] font-medium transition-colors"
						[ngClass]="
							!activeAddIn() && activeTab() === t.id
								? t.id === 'file'
									? 'text-white bg-primary/80 rounded-sm'
									: 'text-foreground after:absolute after:-bottom-px after:left-0 after:right-0 after:h-[2.5px] after:bg-primary'
								: t.id === 'file'
									? 'text-primary hover:bg-primary/15 rounded-sm'
									: 'text-muted-foreground hover:bg-accent/30 hover:text-foreground'
						"
					>
						{{ t.labelKey | translate }}
					</button>
				}
				<!-- The host's add-in tabs (ribbonAddIns), after the fixed tabs. -->
				@for (tab of addInTabs(); track tab.id) {
					<button
						type="button"
						role="tab"
						[attr.data-ribbon-add-in-tab]="tab.id"
						[attr.aria-selected]="activeAddIn() === tab.id"
						[title]="addInTip(tab.label)"
						(click)="selectAddIn.emit(tab.id)"
						(dblclick)="tabDoubleClick(tab.id)"
						class="relative whitespace-nowrap px-3.5 py-2 text-[12px] font-medium transition-colors"
						[ngClass]="
							activeAddIn() === tab.id
								? 'text-foreground after:absolute after:-bottom-px after:left-0 after:right-0 after:h-[2.5px] after:bg-primary'
								: 'text-muted-foreground hover:bg-accent/30 hover:text-foreground'
						"
					>
						{{ tab.label }}
					</button>
				}
				<!-- Contextual tabs (Shape Format, ...): shown for the selection, never auto-selected. -->
				@for (id of contextualTabs(); track id) {
					<button
						type="button"
						role="tab"
						[attr.data-ribbon-contextual-tab]="id"
						[attr.aria-selected]="!activeAddIn() && activeTab() === id"
						[title]="tabTip(contextualLabelKey(id))"
						(click)="selectTab.emit(id)"
						(dblclick)="tabDoubleClick(id)"
						class="relative whitespace-nowrap px-3.5 py-2 text-[12px] font-medium text-primary transition-colors"
						[ngClass]="
							!activeAddIn() && activeTab() === id
								? 'after:absolute after:-bottom-px after:left-0 after:right-0 after:h-[2.5px] after:bg-primary'
								: 'hover:bg-primary/10'
						"
					>
						{{ contextualLabelKey(id) | translate }}
					</button>
				}
			</div>

			<!-- Tab-row right actions (Record, Comments, Share), mirroring React's TabRowActions -->
			<div class="flex shrink-0 items-center gap-1 pr-1">
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
				<!-- The live region announces a connected session (Angular has no separate
				     status pill on this row). -->
				<div
					role="status"
					class="flex items-center"
					[attr.aria-label]="
						collabConnected()
							? ('pptx.collaboration.statusAriaLabel'
								| translate: { status: 'pptx.collaboration.status.connected' | translate })
							: null
					"
				>
					<pptx-ui-ribbon-actions
						data-pptx-chrome="tab-row-actions"
						[state]="actionsState()"
						(comments-toggle)="toggleComments.emit()"
						(share-request)="share.emit()"
					></pptx-ui-ribbon-actions>
				</div>
			</div>

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
	protected readonly contextualLabelKey = contextualTabLabelKey;
	protected readonly visibleTabs = computed(() =>
		filterVisibleTabs(TOOLBAR_TABS, this.hiddenActions()),
	);

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

	/** Double click on a tab (not File) switches the ribbon between docked and tabs only. */
	protected tabDoubleClick(tab: string): void {
		if (tab !== 'file') {
			this.displayController()?.toggleCollapsed();
		}
	}

	/** ScreenTip-styled tooltip of a host tab: its label is the host's own text, not a key. */
	protected addInTip(label: string): string | null {
		return this.viewerOpts ? (this.viewerOpts.screenTip(label) ?? null) : label;
	}

	/** ScreenTip-styled tab tooltip (null suppresses the title attribute). */
	protected tabTip(labelKey: string): string | null {
		const label = this.translate.instant(labelKey) as string;
		return this.viewerOpts ? (this.viewerOpts.screenTip(label) ?? null) : label;
	}
}
