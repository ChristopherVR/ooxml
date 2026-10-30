/**
 * custom-shows.component.ts: Slide Show > Custom Slide Show > Custom Shows dialog.
 *
 * Selector: `pptx-custom-shows`
 *
 * New, Copy and Remove for the shows and, for the selected one, PowerPoint's two lists (slides in
 * the presentation, slides in the custom show) with Add, Remove and Move Up/Down. Every decision
 * is the shared `createCustomShowsController`; this component only draws it. The host renders it
 * with `@if`, so each open starts from a fresh draft of the deck's shows.
 */

import {
	ChangeDetectionStrategy,
	Component,
	computed,
	DestroyRef,
	inject,
	input,
	output,
	signal,
	CUSTOM_ELEMENTS_SCHEMA,
} from '@angular/core';
import type { OnInit, Signal } from '@angular/core';
import { LucideX } from '@lucide/angular';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import { createCustomShowsController } from 'ooxml-ui/pptx';
import type { CustomShowsController, CustomShowsState } from 'ooxml-ui/pptx';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';

import { controllerSignal } from './controller-signal';

/** What OK commits. */
export interface CustomShowsSave {
	shows: PptxCustomShow[];
	/** Slides with jumps into removed shows cleared, or undefined when none changed. */
	clearedSlides: PptxSlide[] | undefined;
}

@Component({
	selector: 'pptx-custom-shows',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	imports: [TranslatePipe, LucideX],
	templateUrl: './custom-shows.component.html',
	styleUrl: './custom-shows.component.css',
})
export class CustomShowsComponent implements OnInit {
	private readonly translate = inject(TranslateService);
	private readonly destroyRef = inject(DestroyRef);

	/** All slides in the presentation. */
	readonly slides = input<readonly PptxSlide[]>([]);
	/** Existing custom shows (the dialog edits a draft of them). */
	readonly customShows = input<readonly PptxCustomShow[]>([]);
	/** Id of the show playback is restricted to, or null. */
	readonly activeCustomShowId = input<string | null>(null);

	/** OK: the shows to keep, plus the slides with jumps into removed shows cleared. */
	readonly save = output<CustomShowsSave>();
	/** Restrict (or, with null, stop restricting) the slide show to one show. */
	readonly setActive = output<string | null>();
	/** Show: the dialog committed; start the show with this id. */
	readonly run = output<string>();
	/** The dialog should close. */
	readonly close = output<void>();

	/** The id the picker shows; "" is All Slides. */
	protected readonly active = signal<string>('');

	private controller: CustomShowsController | undefined;
	private readonly snapshot = signal<Signal<CustomShowsState> | null>(null);

	/** The dialog's view model, redrawn on every controller change. */
	protected readonly view = computed(() => {
		const snapshot = this.snapshot();
		if (!snapshot || !this.controller) {
			return null;
		}
		snapshot();
		return this.controller.view(this.slides());
	});

	/** The picker value, re-derived so a removed show cannot stay selected. */
	protected readonly pickerValue = computed(() => {
		const view = this.view();
		const active = this.active();
		return view?.shows.some(({ id }) => id === active) ? active : '';
	});

	protected get ctl(): CustomShowsController {
		return this.controller as CustomShowsController;
	}

	ngOnInit(): void {
		this.active.set(this.activeCustomShowId() ?? '');
		this.controller = createCustomShowsController(this.customShows(), this.activeCustomShowId());
		this.snapshot.set(controllerSignal(this.controller, this.destroyRef));
	}

	protected newShowName(): string {
		return this.translate.instant('pptx.customShows.newDefaultName');
	}

	protected onPick(event: Event): void {
		const next = (event.target as HTMLSelectElement).value;
		this.active.set(next);
		this.setActive.emit(next || null);
	}

	protected onRename(event: Event): void {
		this.ctl.rename((event.target as HTMLInputElement).value);
	}

	private commit(): PptxCustomShow[] {
		const result = this.ctl.result(this.slides());
		this.save.emit({ shows: result.shows, clearedSlides: result.slides });
		// A show removed in this session must not leave playback pinned to an id nothing resolves.
		this.setActive.emit(result.shows.some(({ id }) => id === this.active()) ? this.active() : null);
		return result.shows;
	}

	protected ok(): void {
		this.commit();
		this.close.emit();
	}

	protected showSelected(): void {
		const id = this.ctl.getSnapshot().selectedId;
		const show = this.commit().find((candidate) => candidate.id === id);
		if (show) {
			this.run.emit(show.id);
			this.close.emit();
		}
	}
}
