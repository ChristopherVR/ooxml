/**
 * find-replace-bar.component.ts: Floating find-and-replace bar.
 *
 * Selector: `pptx-find-replace-bar`
 *
 * Ported from:
 *   packages/react/src/viewer/components/FindReplacePanel.tsx
 *
 * This component is **purely presentational**: it holds no slide data, no
 * search state and performs no text mutation itself. Its query, replacement
 * and scopes are inputs from the viewer's Find and Replace session (so a
 * reopened bar shows what the session kept); the bar shows the user's own
 * edit at once, before the session hands it back, and every edit goes out:
 *
 *   - `find` when the query or a scope changes, `replacementChange` when the
 *     replacement does.
 *   - `navigate`, `replaceOne`, `replaceAll` and `close` for the buttons.
 *
 * Usage:
 * ```html
 * <pptx-find-replace-bar
 *   [query]="session().query"
 *   [replacement]="session().replacement"
 *   [matchCount]="totalMatches()"
 *   [matchIndex]="activeResultIndex()"
 *   (find)="onFind($event)"
 *   (replacementChange)="onReplacementChange($event)"
 *   (navigate)="onNavigate($event)"
 *   (replaceOne)="onReplaceOne($event)"
 *   (replaceAll)="onReplaceAll($event)"
 *   (close)="closeFindReplace()"
 * />
 * ```
 */

import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	ElementRef,
	HostListener,
	computed,
	input,
	linkedSignal,
	output,
	viewChild,
} from '@angular/core';
import type { WritableSignal } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { findReplacementAvailability } from 'ooxml-ui/pptx';
import type { FindResult } from 'ooxml-ui/pptx';

// ---------------------------------------------------------------------------
// Event payload types
// ---------------------------------------------------------------------------

/** Emitted when the user changes the find query or the case-sensitive toggle. */
export interface FindEvent {
	query: string;
	matchCase: boolean;
	includeNotes?: boolean;
	includeChartTitles?: boolean;
	includeChartData?: boolean;
}

/** Emitted when the user confirms a replacement action. */
export interface ReplaceEvent {
	query: string;
	replacement: string;
	matchCase: boolean;
	includeNotes?: boolean;
	includeChartTitles?: boolean;
	includeChartData?: boolean;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-find-replace-bar',
	standalone: true,
	imports: [TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div
			class="pptx-frb"
			role="dialog"
			[attr.aria-label]="'pptx.findReplace.ariaLabel' | translate"
		>
			<!-- ── Find row ───────────────────────────────────────────────── -->
			<div class="pptx-frb__row">
				<label
					><pptx-ui-checkbox
						[attr.aria-label]="'pptx.findReplace.includeChartTitles' | translate"
						[checked]="shown.includeChartTitles()"
						(change)="toggleIncludeChartTitles()"
					></pptx-ui-checkbox
					>{{ 'pptx.findReplace.includeChartTitles' | translate }}</label
				>
				<label
					><pptx-ui-checkbox
						[attr.aria-label]="'pptx.findReplace.includeChartData' | translate"
						[checked]="shown.includeChartData()"
						(change)="toggleIncludeChartData()"
					></pptx-ui-checkbox
					>{{ 'pptx.findReplace.includeChartData' | translate }}</label
				>
				<label
					><pptx-ui-checkbox
						[attr.aria-label]="'pptx.findReplace.includeNotes' | translate"
						[checked]="shown.includeNotes()"
						(change)="toggleIncludeNotes()"
					></pptx-ui-checkbox
					>{{ 'pptx.findReplace.includeNotes' | translate }}</label
				>
				<input
					#findInput
					class="pptx-frb__input"
					type="search"
					[placeholder]="'pptx.findReplace.findPlaceholder' | translate"
					[attr.aria-label]="'pptx.findReplace.searchText' | translate"
					[value]="shown.query()"
					(input)="onQueryInput($event)"
					(keydown.enter)="navigate.emit(1)"
				/>

				<!-- Case-sensitive toggle -->
				<button
					type="button"
					class="pptx-frb__btn"
					[class.pptx-frb__btn--active]="shown.matchCase()"
					[attr.aria-label]="'pptx.findReplace.matchCase' | translate"
					[attr.aria-pressed]="shown.matchCase()"
					(click)="toggleCase()"
					[title]="'pptx.findReplace.matchCase' | translate"
				>
					Aa
				</button>

				<!-- Match counter -->
				<span class="pptx-frb__count" aria-live="polite" aria-atomic="true">
					@if (shown.query().trim()) {
						@if (matchCount() === 0) {
							{{ 'pptx.findReplace.noMatches' | translate }}
						} @else {
							{{
								'pptx.findReplace.matchCount'
									| translate: { current: displayIndex(), total: matchCount() }
							}}
						}
					}
				</span>

				<button
					type="button"
					class="pptx-frb__btn"
					[attr.aria-label]="'pptx.findReplace.previousMatch' | translate"
					[disabled]="matchCount() === 0"
					(click)="navigate.emit(-1)"
				>
					&#8593;
				</button>

				<button
					type="button"
					class="pptx-frb__btn"
					[attr.aria-label]="'pptx.findReplace.nextMatch' | translate"
					[disabled]="matchCount() === 0"
					(click)="navigate.emit(1)"
				>
					&#8595;
				</button>

				<button
					type="button"
					class="pptx-frb__btn pptx-frb__btn--close"
					[attr.aria-label]="'pptx.findReplace.closeAriaLabel' | translate"
					(click)="close.emit()"
				>
					&#10005;
				</button>
			</div>

			<!-- ── Replace row ────────────────────────────────────────────── -->
			<div class="pptx-frb__row pptx-frb__row--replace">
				<input
					class="pptx-frb__input"
					type="text"
					[placeholder]="'pptx.findReplace.replacePlaceholder' | translate"
					[attr.aria-label]="'pptx.findReplace.replacementText' | translate"
					[value]="shown.replacement()"
					(input)="onReplacementInput($event)"
					(keydown.enter)="emitReplaceOne()"
				/>

				<button
					type="button"
					class="pptx-frb__action-btn"
					[disabled]="!availability().current"
					(click)="emitReplaceOne()"
					[attr.aria-label]="'pptx.findReplace.replaceCurrent' | translate"
				>
					{{ 'pptx.findReplace.replace' | translate }}
				</button>

				<button
					type="button"
					class="pptx-frb__action-btn"
					[disabled]="!availability().all"
					(click)="emitReplaceAll()"
					[attr.aria-label]="'pptx.findReplace.replaceAllMatches' | translate"
				>
					{{ 'pptx.findReplace.replaceAll' | translate }}
				</button>
			</div>
			@if (availability().hasReadOnly) {
				<p role="status">
					{{
						(shown.includeChartData()
							? 'pptx.findReplace.readOnlyChartText'
							: 'pptx.findReplace.readOnlyChartTitle'
						) | translate
					}}
				</p>
			}
		</div>
	`,
	styles: [
		`
			:host {
				display: block;
				position: fixed;
				top: 3.5rem;
				right: 1rem;
				z-index: 100;
			}

			.pptx-frb {
				display: flex;
				flex-direction: column;
				gap: 0.375rem;
				min-width: 26rem;
				padding: 0.5rem 0.625rem;
				border: 1px solid rgba(255, 255, 255, 0.12);
				border-radius: 0.375rem;
				background: #1e1e1e;
				color: #e5e5e5;
				box-shadow: 0 8px 24px rgba(0, 0, 0, 0.5);
				font-size: 0.8125rem;
			}

			.pptx-frb__row {
				display: flex;
				align-items: center;
				gap: 0.375rem;
			}

			.pptx-frb__row--replace {
				padding-top: 0.125rem;
				border-top: 1px solid rgba(255, 255, 255, 0.07);
			}

			.pptx-frb__input {
				flex: 1;
				min-width: 0;
				padding: 0.3rem 0.5rem;
				border: 1px solid rgba(255, 255, 255, 0.15);
				border-radius: 0.25rem;
				background: rgba(255, 255, 255, 0.06);
				color: inherit;
				font-size: inherit;
				outline: none;
			}

			.pptx-frb__input:focus {
				border-color: #3b82f6;
				background: rgba(59, 130, 246, 0.08);
			}

			/* Remove default webkit search cancel button */
			.pptx-frb__input::-webkit-search-cancel-button {
				display: none;
			}

			.pptx-frb__count {
				white-space: nowrap;
				color: rgba(255, 255, 255, 0.5);
				font-size: 0.75rem;
				min-width: 5.5rem;
				text-align: right;
			}

			.pptx-frb__btn {
				display: inline-flex;
				align-items: center;
				justify-content: center;
				width: 1.75rem;
				height: 1.75rem;
				padding: 0;
				border: 1px solid rgba(255, 255, 255, 0.1);
				border-radius: 0.25rem;
				background: rgba(255, 255, 255, 0.06);
				color: inherit;
				cursor: pointer;
				transition: background 0.12s;
				flex-shrink: 0;
				font-size: 0.875rem;
				line-height: 1;
			}

			.pptx-frb__btn:hover:not(:disabled) {
				background: rgba(255, 255, 255, 0.14);
			}

			.pptx-frb__btn:disabled {
				opacity: 0.35;
				cursor: not-allowed;
			}

			.pptx-frb__btn--active {
				background: rgba(59, 130, 246, 0.3);
				border-color: rgba(59, 130, 246, 0.6);
				color: #93c5fd;
			}

			.pptx-frb__btn--close {
				border-color: transparent;
				font-size: 0.75rem;
			}

			.pptx-frb__action-btn {
				padding: 0.25rem 0.625rem;
				border: 1px solid rgba(255, 255, 255, 0.15);
				border-radius: 0.25rem;
				background: rgba(255, 255, 255, 0.08);
				color: inherit;
				font-size: 0.75rem;
				cursor: pointer;
				white-space: nowrap;
				transition: background 0.12s;
				flex-shrink: 0;
			}

			.pptx-frb__action-btn:hover:not(:disabled) {
				background: rgba(255, 255, 255, 0.16);
			}

			.pptx-frb__action-btn:disabled {
				opacity: 0.35;
				cursor: not-allowed;
			}
		`,
	],
})
export class FindReplaceBarComponent {
	// -------------------------------------------------------------------------
	// Inputs
	// -------------------------------------------------------------------------

	/**
	 * Total number of matches found across all slides for the current query.
	 * Kept at 0 when no search has been performed yet.
	 */
	readonly matchCount = input<number>(0);

	/**
	 * Zero-based index of the currently highlighted match.
	 * Used to derive the 1-based display counter.
	 */
	readonly matchIndex = input<number>(-1);
	readonly matches = input<readonly FindResult[] | undefined>();
	readonly availability = computed(() =>
		this.matches()
			? findReplacementAvailability(this.matches()!, this.matchIndex())
			: { current: this.matchCount() > 0, all: this.matchCount() > 0, hasReadOnly: false },
	);

	// -------------------------------------------------------------------------
	// Outputs
	// -------------------------------------------------------------------------

	/**
	 * Emitted whenever the find query or a scope toggle changes.
	 * The parent searches again and updates the inputs.
	 */
	readonly find = output<FindEvent>();

	/** Emitted as the replacement text is typed. */
	readonly replacementChange = output<string>();

	/**
	 * Emitted when the user clicks ↑ / ↓ or presses Enter in the find input.
	 * Payload is `1` (next) or `-1` (previous).
	 * The parent advances `matchIndex` and navigates to the matching slide.
	 */
	readonly navigate = output<1 | -1>();

	/**
	 * Emitted when the user clicks "Replace" (single match).
	 * The parent calls `replaceMatch(...)` on `EditorStateService`.
	 */
	readonly replaceOne = output<ReplaceEvent>();

	/**
	 * Emitted when the user clicks "Replace All".
	 * The parent calls `replaceInSlides(...)` and commits the result to history.
	 */
	readonly replaceAll = output<ReplaceEvent>();

	/** Emitted when the user closes the bar (Escape or ✕ button). */
	readonly close = output<void>();

	// -------------------------------------------------------------------------
	// Template reference
	// -------------------------------------------------------------------------

	private readonly findInputRef = viewChild<ElementRef<HTMLInputElement>>('findInput');

	// -------------------------------------------------------------------------
	// The session's fields (inputs; edits go out through `find` and `replacementChange`)
	// -------------------------------------------------------------------------

	/** Current text in the find input. */
	readonly query = input('');

	/** Current text in the replacement input. */
	readonly replacement = input('');

	/** Whether the search should be case-sensitive. */
	readonly matchCase = input(false);
	readonly includeNotes = input(false);
	readonly includeChartTitles = input(true);
	readonly includeChartData = input(false);

	/**
	 * The fields as the bar shows and sends them: each follows its input, and the user's own edit
	 * applies at once. Zoneless change detection hands the bar the session's values on its next
	 * pass, so an edit composed from the inputs alone could undo the edit just before it (a scope
	 * ticked, then a query typed before that pass).
	 */
	protected readonly shown = {
		query: linkedSignal(() => this.query()),
		replacement: linkedSignal(() => this.replacement()),
		matchCase: linkedSignal(() => this.matchCase()),
		includeNotes: linkedSignal(() => this.includeNotes()),
		includeChartTitles: linkedSignal(() => this.includeChartTitles()),
		includeChartData: linkedSignal(() => this.includeChartData()),
	};

	// -------------------------------------------------------------------------
	// Derived
	// -------------------------------------------------------------------------

	/** 1-based display index for the counter (e.g. "2 / 5"). Clamps to bounds. */
	readonly displayIndex = computed<number>(() => {
		const count = this.matchCount();
		if (count === 0) {
			return 0;
		}
		const idx = this.matchIndex();
		if (idx < 0) {
			return 1;
		}
		return Math.min(idx + 1, count);
	});

	// -------------------------------------------------------------------------
	// Keyboard handling
	// -------------------------------------------------------------------------

	/** Escape anywhere on the document closes the bar. */
	@HostListener('document:keydown', ['$event'])
	onDocumentKeydown(event: KeyboardEvent): void {
		if (event.key === 'Escape') {
			event.preventDefault();
			this.close.emit();
		}
	}

	// -------------------------------------------------------------------------
	// Event handlers
	// -------------------------------------------------------------------------

	/** The query and scopes as the bar shows them. */
	private findEvent(): FindEvent {
		const shown = this.shown;
		return {
			query: shown.query(),
			matchCase: shown.matchCase(),
			includeNotes: shown.includeNotes(),
			includeChartTitles: shown.includeChartTitles(),
			includeChartData: shown.includeChartData(),
		};
	}

	/** Flip a scope where the bar shows it, then search with it. */
	private toggle(scope: WritableSignal<boolean>): void {
		scope.update((on) => !on);
		this.find.emit(this.findEvent());
	}

	onQueryInput(event: Event): void {
		this.shown.query.set((event.target as HTMLInputElement).value);
		this.find.emit(this.findEvent());
	}

	onReplacementInput(event: Event): void {
		const replacement = (event.target as HTMLInputElement).value;
		this.shown.replacement.set(replacement);
		this.replacementChange.emit(replacement);
	}

	toggleCase(): void {
		this.toggle(this.shown.matchCase);
	}

	toggleIncludeNotes(): void {
		this.toggle(this.shown.includeNotes);
	}
	toggleIncludeChartTitles(): void {
		this.toggle(this.shown.includeChartTitles);
	}
	toggleIncludeChartData(): void {
		this.toggle(this.shown.includeChartData);
	}
	emitReplaceOne(): void {
		if (!this.availability().current) {
			return;
		}
		this.replaceOne.emit({ ...this.findEvent(), replacement: this.shown.replacement() });
	}

	emitReplaceAll(): void {
		if (!this.availability().all) {
			return;
		}
		this.replaceAll.emit({ ...this.findEvent(), replacement: this.shown.replacement() });
	}

	/** Focus the find input (called by the parent after toggling the bar open). */
	focusFindInput(): void {
		this.findInputRef()?.nativeElement.focus();
	}
}
