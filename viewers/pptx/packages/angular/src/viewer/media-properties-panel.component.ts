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
	mediaTrimEndAbsoluteMs,
	mediaTrimEndMsFromAbsoluteMs,
	mediaTrimResetPatch,
	mediaTrimResetState,
	MEDIA_PLAYBACK_SECTIONS,
} from 'ooxml-ui/pptx';
import type { MediaBookmark, MediaPptxElement, PptxElement } from 'pptx-viewer-core';

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
			<div class="grid two">
				<label>
					<span>{{ 'pptx.media.trimStartTime' | translate }}</span>
					<input
						type="number"
						min="0"
						[value]="media().trimStartMs ?? 0"
						(change)="numberPatch('trimStartMs', $event)"
					/>
				</label>
				<label>
					<span>{{ 'pptx.media.trimEndTime' | translate }}</span>
					<input
						type="number"
						min="0"
						[value]="trimEndAbsoluteMs()"
						(change)="trimEndPatch($event)"
					/>
				</label>
			</div>
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
							type="number"
							min="0"
							step="0.1"
							[value]="bookmark.time"
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
	readonly patch = output<Partial<PptxElement>>();

	private readonly loader = inject(LoadContentService, { optional: true });
	protected readonly media = computed(() => this.element());
	protected readonly mediaDataUrls = computed(() => this.loader?.mediaDataUrls() ?? new Map());
	protected readonly playbackSections = MEDIA_PLAYBACK_SECTIONS;
	/**
	 * `trimEndMs` is `p14:trim/@end`, the distance from the clip's TAIL
	 * (COM-verified), not an absolute stop time. The "Trim end" field shows the
	 * user an absolute position and converts back on commit, mirroring React's
	 * `MediaInspector` / Vue's `MediaPropertiesPanel.vue` via the shared
	 * `media-trim-range.ts` (this raw ms field used to bind the tail distance
	 * directly, so typing "the last 5s" required computing duration-minus-5000
	 * by hand).
	 */
	private readonly durationMs = computed(() => (this.media().metadata?.duration ?? 0) * 1000);
	protected readonly trimEndAbsoluteMs = computed(() =>
		mediaTrimEndAbsoluteMs(this.durationMs(), this.media().trimEndMs ?? 0),
	);
	protected numberPatch(key: keyof MediaPptxElement, event: Event): void {
		this.patch.emit({
			[key]: Number((event.target as HTMLInputElement).value),
		} as Partial<PptxElement>);
	}

	/** Convert the typed absolute end position back to `trimEndMs`'s tail distance. */
	protected trimEndPatch(event: Event): void {
		const absoluteMs = Number((event.target as HTMLInputElement).value);
		this.patch.emit({
			trimEndMs: mediaTrimEndMsFromAbsoluteMs(this.durationMs(), absoluteMs),
		} as Partial<PptxElement>);
	}

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

	protected updateBookmark(index: number, key: keyof MediaBookmark, event: Event): void {
		const target = event.target as HTMLInputElement;
		const bookmarks = (this.media().bookmarks ?? []).map((bookmark, i) =>
			i === index
				? { ...bookmark, [key]: key === 'time' ? Number(target.value) : target.value }
				: bookmark,
		);
		this.patch.emit({ bookmarks } as Partial<PptxElement>);
	}
}
