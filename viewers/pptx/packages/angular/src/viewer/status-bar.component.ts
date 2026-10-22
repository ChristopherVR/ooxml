/**
 * status-bar.component.ts: bottom status bar for the Angular editor chrome.
 *
 * A thin adapter around the shared `pptx-ui-status-bar` element: the slide
 * counter, save indicator, notes toggle, view-mode buttons and zoom cluster are
 * rendered by the shared view. This component maps viewer state onto the
 * element's controlled state and re-emits its typed intents as outputs the
 * {@link PowerPointViewerComponent} already has handlers for.
 *
 * A host may project a connection indicator with `[pptxCollabStatus]`; the
 * projected element must also carry `slot="collaboration"`.
 */
import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
	output,
} from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { TranslateService } from '@ngx-translate/core';
import { dispatchStatusBarIntent, statusBarState } from 'ooxml-ui/pptx';
import type { StatusBarIntentHandlers, StatusBarViewState, ToolbarActionId } from 'ooxml-ui/pptx';
import { map, merge, startWith } from 'rxjs';

import type { AutosaveStatus } from './autosave.service';

@Component({
	selector: 'pptx-status-bar',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	host: { class: 'contents' },
	template: `
		<pptx-ui-status-bar [state]="view()" (status-request)="request($event)">
			<ng-content select="[pptxCollabStatus]"></ng-content>
		</pptx-ui-status-bar>
	`,
})
export class StatusBarComponent {
	readonly slideIndex = input<number>(0);
	readonly slideCount = input<number>(0);
	readonly canEdit = input<boolean>(false);
	readonly dirty = input<boolean>(false);
	/** Current autosave engine status; drives the save-state text + colour. */
	readonly autosaveStatus = input<AutosaveStatus | undefined>(undefined);
	readonly notesOpen = input<boolean>(false);
	readonly zoomPercent = input<number>(100);
	/** True when the slide-sorter overlay is open (active-state styling). */
	readonly sorterActive = input<boolean>(false);
	/** True when the presentation overlay is open (active-state styling). */
	readonly presenting = input<boolean>(false);
	/** Toolbar buttons the host wants hidden (notes/fullscreen/zoom independently). */
	readonly hiddenActions = input<ToolbarActionId[]>([]);

	readonly toggleNotes = output<void>();
	readonly normalView = output<void>();
	readonly openSorter = output<void>();
	readonly slideShow = output<void>();
	readonly zoomIn = output<void>();
	readonly zoomOut = output<void>();
	readonly zoomReset = output<void>();

	private readonly translate = inject(TranslateService);
	/** Changes on language/dictionary updates so OnPush re-translates the view. */
	private readonly translations = toSignal(
		merge(this.translate.onLangChange, this.translate.onTranslationChange).pipe(
			map(() => Date.now()),
			startWith(0),
		),
		{ initialValue: 0 },
	);

	private readonly intents: StatusBarIntentHandlers = {
		notes: () => this.toggleNotes.emit(),
		normal: () => this.normalView.emit(),
		sorter: () => this.openSorter.emit(),
		slideShow: () => this.slideShow.emit(),
		zoomOut: () => this.zoomOut.emit(),
		zoomFit: () => this.zoomReset.emit(),
		zoomIn: () => this.zoomIn.emit(),
	};

	protected view(): StatusBarViewState {
		this.translations();
		return statusBarState({
			slideCount: this.slideCount(),
			activeSlideIndex: this.slideIndex(),
			dirty: this.dirty(),
			autosave: this.autosaveStatus(),
			zoomPercent: this.zoomPercent(),
			notesExpanded: this.notesOpen(),
			viewMode: this.presenting() ? 'slideShow' : this.sorterActive() ? 'sorter' : 'normal',
			hiddenActions: this.hiddenActions(),
			translate: (key, params) => this.translate.instant(key, params),
		});
	}

	protected request(event: Event): void {
		dispatchStatusBarIntent(event, this.intents);
	}
}
