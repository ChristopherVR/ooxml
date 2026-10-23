/**
 * autosave-recovery.service.ts: the Angular adapter of the shared crash-recovery
 * offer (`createAutosaveRecoveryController` in `ooxml-ui/pptx`).
 *
 * Provide it once on the viewer component. The component wires its accessors
 * via {@link bind}; the resulting {@link prompt} signal feeds the dialog.
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
import {
	createAutosaveRecoveryController,
	discardAutosaveRecovery,
	probeAutosaveRecovery,
} from 'ooxml-ui/pptx';
import type { AutosaveRecord, DocumentStorage } from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';

/** Live host accessors the recovery probe reads (all reactive). */
export interface AutosaveRecoveryHost {
	/** IndexedDB key of the open deck. Nothing to look up without one. */
	readonly filePath: () => string | undefined;
	/** Public document name rendered in the recovery message. */
	readonly fileName: () => string | undefined;
	/** Host `storage` input: recovery snapshots go to its recovery methods when it has them. */
	readonly storage?: () => DocumentStorage | undefined;
	/** True while the load pipeline is running. */
	readonly loading: () => boolean;
	/** Load error, if any. */
	readonly error: () => string | null | undefined;
	/** Slides currently rendered. */
	readonly slideCount: () => number;
	/**
	 * Whether the host permits autosave at all (`autosave` input not `false`). A
	 * user who merely switched the toggle off is still offered a pre-crash
	 * snapshot; a host that forbade the feature is not.
	 */
	readonly autosaveAllowed: () => boolean;
	/** Load the recovered bytes into the viewer. */
	readonly restore: (bytes: Uint8Array) => void;
	/** Called with the stored record when a snapshot is offered (the host's `autosaveRecovery`). */
	readonly offered?: (record: AutosaveRecord) => void;
}

@Injectable()
export class AutosaveRecoveryService {
	private readonly injector = inject(Injector);
	private readonly destroyRef = inject(DestroyRef);

	private host: AutosaveRecoveryHost | null = null;
	private readonly controller = createAutosaveRecoveryController({
		getStorage: () => this.host?.storage?.(),
		load: (bytes) => this.host?.restore(bytes),
		onOffered: (record) => this.host?.offered?.(record),
		probe: probeAutosaveRecovery,
		discardRecord: discardAutosaveRecovery,
	});
	private readonly state = controllerSignal(this.controller);

	/** What the dialog should say, or null when there is nothing to offer. */
	readonly prompt = computed(() => this.state().prompt);
	/** The destructive delete is still pending. */
	readonly discarding = computed(() => this.state().discarding);

	bind(host: AutosaveRecoveryHost): void {
		this.host = host;
		this.destroyRef.onDestroy(() => this.controller.destroy());
		effect(
			() => {
				const input = {
					filePath: host.filePath(),
					fileName: host.fileName(),
					loading: host.loading(),
					error: host.error() ?? null,
					slideCount: host.slideCount(),
					autosaveAllowed: host.autosaveAllowed(),
				};
				untracked(() => this.controller.check(input));
			},
			{ injector: this.injector },
		);
	}

	/** The user accepted: load the snapshot bytes. */
	restore(): void {
		void this.controller.restore().catch(() => {
			// The load path reports its own error.
		});
	}

	/** The user declined: drop the snapshot. */
	discard(): void {
		void this.controller.discard();
	}
}
