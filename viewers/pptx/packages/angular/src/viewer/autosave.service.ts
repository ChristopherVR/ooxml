/**
 * autosave.service.ts: the Angular adapter of the shared recovery autosave
 * engine (`createAutosaveController` in `ooxml-ui/pptx`), the counterpart of
 * React's `useAutosave` hook.
 *
 * The engine follows PowerPoint's AutoRecover: while there are unsaved edits a
 * recovery snapshot is written every interval (the shared
 * {@link resolveAutosaveIntervalMs}: the host's explicit cadence, else File >
 * Options > Save AutoRecover, else 120s), and nothing is written while the deck
 * is unchanged. This service only feeds it the component's signals and exposes
 * its state as a signal.
 *
 * Provide it once on the viewer component (`providers: [AutosaveService]`). The
 * component wires its reactive accessors via {@link bind}; the resulting
 * {@link status} signal feeds both the title bar and the status bar.
 */

import {
	computed,
	DestroyRef,
	effect,
	inject,
	Injectable,
	Injector,
	untracked,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	createAutosaveController,
	resolveAutosaveIntervalMs,
	saveAutosaveSnapshot,
} from 'ooxml-ui/pptx';
import type {
	AutosaveDisabledReason,
	DocumentStorage,
	InlineTextDraft,
	PptxAutosaveStatus,
} from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';

/** The one autosave status every binding renders (`PptxAutosaveStatus` in `ooxml-ui/pptx`). */
export type AutosaveStatus = PptxAutosaveStatus;

/** Live host accessors the autosave engine reads (all reactive). */
export interface AutosaveHost {
	/**
	 * Whether autosave actually runs: the shared `resolveAutosaveActivation`
	 * verdict (the host's `autosave` input as a ceiling, the title-bar toggle as
	 * the user preference inside it), not the raw toggle.
	 */
	readonly enabled: () => boolean;
	/** File path/name keying the recovery snapshot. Autosave is inert when unset. */
	readonly filePath: () => string | undefined;
	/** Whether the deck has unsaved edits. */
	readonly isDirty: () => boolean;
	/** Host `storage` input: recovery snapshots go to its recovery methods when it has them. */
	readonly storage?: () => DocumentStorage | undefined;
	/** Serialise the current edited deck to `.pptx` bytes (null skips the write). */
	readonly serialize: (draft: InlineTextDraft | null) => Promise<Uint8Array | null>;
	/** The element whose inline editor is open; a closed editor's draft is dropped. */
	readonly inlineEditingId?: () => string | null;
	/**
	 * The resolved cadence in milliseconds (shared `resolveAutosaveIntervalMs`:
	 * host `autosaveIntervalMs` input > File > Options > Save AutoRecover
	 * cadence > 120s default).
	 */
	readonly intervalMs?: () => number;
	/** Why autosave is inactive, straight from the shared activation verdict. */
	readonly disabledReason?: () => AutosaveDisabledReason | undefined;
	/**
	 * The signals a snapshot is built from (reassigned, never mutated). While the
	 * deck is dirty, a change in any of them is an edit that starts the
	 * AutoRecover interval. Omit it and only the clean-to-dirty transition counts.
	 */
	readonly changeSources?: () => readonly unknown[];
	/** Called with the serialised deck after each snapshot is written (the host's `autosaved`). */
	readonly saved?: (data: Uint8Array) => void;
}

@Injectable()
export class AutosaveService {
	private readonly injector = inject(Injector);
	private readonly destroyRef = inject(DestroyRef);
	private readonly translate = inject(TranslateService);

	private host: AutosaveHost | null = null;
	private readonly controller = createAutosaveController({
		getActivation: () =>
			this.host
				? { active: this.host.enabled(), reason: this.host.disabledReason?.() }
				: { active: false },
		getFilePath: () => this.host?.filePath(),
		getIntervalMs: () => resolveAutosaveIntervalMs({ hostIntervalMs: this.host?.intervalMs?.() }),
		serialize: (draft) => this.host?.serialize(draft) ?? Promise.resolve(null),
		getInlineEditingId: () => this.host?.inlineEditingId?.(),
		getStorage: () => this.host?.storage?.(),
		onSaved: (data) => this.host?.saved?.(data),
		errorMessage: () => this.translate.instant('pptx.autosave.error'),
		writeSnapshot: saveAutosaveSnapshot,
	});
	private readonly state = controllerSignal(this.controller);

	/** Current autosave status, surfaced in the title bar and status bar. */
	readonly status = computed(() => this.state().status);

	/**
	 * Wire the host accessors (called once from the component constructor). One
	 * effect re-reads the activation and cadence; another reports edits.
	 */
	bind(host: AutosaveHost): void {
		this.host = host;
		this.destroyRef.onDestroy(() => this.controller.destroy());
		effect(
			() => {
				host.enabled();
				host.filePath();
				host.intervalMs?.();
				host.disabledReason?.();
				untracked(() => this.controller.refresh());
			},
			{ injector: this.injector },
		);
		effect(
			() => {
				const dirty = host.isDirty();
				const sources = host.changeSources?.() ?? [];
				untracked(() => this.controller.observe(dirty, sources));
			},
			{ injector: this.injector },
		);
	}

	/** In-progress inline text: a changed draft is an edit and goes into the snapshot. */
	noteInlineDraft(draft: InlineTextDraft | null): void {
		this.controller.noteInlineDraft(draft);
	}

	/** Manually trigger an autosave right now (mirrors React's `triggerAutosave`). */
	async triggerAutosave(): Promise<void> {
		await this.controller.saveNow();
	}
}
