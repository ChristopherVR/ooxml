import {
	ChangeDetectionStrategy,
	Component,
	computed,
	inject,
	input,
	output,
	CUSTOM_ELEMENTS_SCHEMA,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	MEDIA_TRIM_SECTIONS,
	mediaTrimResetPatch,
	mediaTrimResetState,
	MEDIA_PLAYBACK_SECTIONS,
	formatMediaBookmarkTime,
	parseMediaBookmarkTime,
} from 'ooxml-ui/pptx';
import type { MediaPptxElement, PptxElement } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { LoadContentService } from './load-content.service';
import { MediaPreviewComponent } from './media-preview.component';
import { appendMediaBookmark } from './media-properties-helpers';

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-media-properties-panel',
	standalone: true,
	imports: [TranslatePipe, InspectorFieldsComponent, MediaPreviewComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-media-properties">
			<pptx-media-preview
				[element]="media()"
				[mediaDataUrls]="mediaDataUrls()"
				[canEdit]="canEdit()"
				(patch)="patch.emit($event)"
			/>
			<pptx-inspector-fields
				[sections]="trimSections"
				[target]="trimTarget()"
				[disabled]="!canEdit()"
				(patch)="patch.emit($event)"
			/>
			@if (resetTrimState().visible) {
				<button type="button" class="reset-trim" (click)="patch.emit(trimResetPatch())">
					{{ 'pptx.media.resetTrim' | translate }}
				</button>
			}
			<pptx-inspector-fields
				[sections]="playbackSections"
				[target]="media()"
				[disabled]="!canEdit()"
				(patch)="patch.emit($event)"
			/>
			<section class="bookmarks">
				<div class="bookmark-head">
					<strong>{{ 'pptx.media.bookmarks' | translate }}</strong>
					<button type="button" (click)="addBookmark()">+</button>
				</div>
				@for (bookmark of media().bookmarks ?? []; track $index) {
					<div class="bookmark-row">
						<input
							type="text"
							[value]="bookmark.label"
							(change)="updateBookmark($index, 'label', $event)"
						/>
						<input
							type="text"
							[value]="bookmarkTime(bookmark.time)"
							(change)="updateBookmark($index, 'time', $event)"
						/>
						<button type="button" (click)="removeBookmark($index)">×</button>
					</div>
				}
			</section>
		</div>
	`,
	styles: `
		.pptx-ng-media-properties {
			display: grid;
			gap: 9px;
			font-size: 11px;
		}
		.grid {
			display: grid;
			gap: 6px;
		}
		.two {
			grid-template-columns: 1fr 1fr;
		}
		label {
			display: grid;
			gap: 3px;
			color: var(--pptx-inspector-muted, #aaa);
		}
		.row {
			grid-template-columns: 90px minmax(0, 1fr) auto;
			align-items: center;
		}
		input,
		select {
			box-sizing: border-box;
			min-width: 0;
			width: 100%;
			padding: 4px 6px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
		}
		input[type='range'] {
			padding: 0;
			accent-color: var(--pptx-primary, #2563eb);
		}
		pptx-ui-checkbox {
			width: auto;
			justify-self: end;
		}
		output {
			width: 32px;
			text-align: right;
			font-variant-numeric: tabular-nums;
		}
		.bookmarks {
			display: grid;
			gap: 5px;
			padding-top: 5px;
			border-top: 1px solid var(--pptx-inspector-border, #444);
		}
		.bookmark-head,
		.bookmark-row {
			display: grid;
			grid-template-columns: 1fr auto;
			align-items: center;
			gap: 5px;
		}
		.bookmark-row {
			grid-template-columns: 1fr 76px auto;
		}
		button {
			padding: 3px 7px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
			cursor: pointer;
		}
	`,
})
export class MediaPropertiesPanelComponent {
	readonly element = input.required<MediaPptxElement>();
	readonly canEdit = input<boolean>(true);
	protected readonly resetTrimState = computed(() =>
		mediaTrimResetState(this.media(), this.canEdit()),
	);
	protected readonly trimResetPatch = mediaTrimResetPatch;
	protected readonly bookmarkTime = formatMediaBookmarkTime;
	readonly patch = output<Partial<PptxElement>>();

	private readonly loader = inject(LoadContentService, { optional: true });
	protected readonly media = computed(() => this.element());
	protected readonly mediaDataUrls = computed(() => this.loader?.mediaDataUrls() ?? new Map());
	protected readonly playbackSections = MEDIA_PLAYBACK_SECTIONS;
	/** The clip with its duration: what the shared trim fields are bounded by. */
	private readonly durationMs = computed(() => (this.media().metadata?.duration ?? 0) * 1000);
	protected readonly trimTarget = computed(() => ({
		media: this.media(),
		durationMs: this.durationMs(),
	}));
	protected readonly trimSections = MEDIA_TRIM_SECTIONS;
	protected addBookmark(): void {
		const current = this.media().bookmarks ?? [];
		const bookmarks = appendMediaBookmark(
			current,
			this.media().trimStartMs ?? 0,
			`bookmark-${Date.now()}-${current.length}`,
		);
		this.patch.emit({ bookmarks } as Partial<PptxElement>);
	}

	protected removeBookmark(index: number): void {
		this.patch.emit({
			bookmarks: (this.media().bookmarks ?? []).filter((_, i) => i !== index),
		} as Partial<PptxElement>);
	}

	protected updateBookmark(index: number, key: 'label' | 'time', event: Event): void {
		const target = event.target;
		const bookmarks = this.media().bookmarks ?? [];
		const bookmark = bookmarks[index];
		if (!(target instanceof HTMLInputElement) || !bookmark) {
			return;
		}
		const next = { ...bookmark };
		if (key === 'time') {
			next.time = parseMediaBookmarkTime(target.value, bookmark.time) ?? bookmark.time;
			target.value = formatMediaBookmarkTime(next.time);
		} else {
			next.label = target.value;
		}
		this.patch.emit({ bookmarks: bookmarks.map((item, i) => (i === index ? next : item)) });
	}
}
