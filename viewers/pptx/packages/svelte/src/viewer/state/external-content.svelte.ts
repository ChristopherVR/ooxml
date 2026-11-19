import { resolveExternalContentView } from 'ooxml-ui/pptx';
import type {
	ExternalContentOptions,
	ExternalContentView,
	ViewerExternalContentBar,
} from 'ooxml-ui/pptx';

export interface ExternalContentDeps {
	/** The loader's `loadCount`: each finished load is another presentation. */
	getLoadCount: () => number;
	/** The host `allowExternalContent` prop. */
	getHost: () => boolean | undefined;
	/** The viewer options; Trust Center > "Allow external content" applies at once. */
	getOptions: () => ExternalContentOptions;
	/** What the viewer draws: its slides, template elements and table style map. */
	getContent: () => unknown;
}

interface Consent {
	load: number;
	enabled: boolean;
	closed: boolean;
}

/**
 * PowerPoint's security bar for linked pictures and media, and whether this
 * viewer draws them (shared `resolveExternalContentView`). Consent is this
 * viewer's own and belongs to the presentation it was given for: the next load
 * asks again, while slides adopted from a collaboration room keep it. Enable
 * Content only redraws (`provideLinkedContent`): the deck, undo history, dirty
 * state, autosave and a collaboration session carry on untouched.
 */
export class ExternalContentState implements ViewerExternalContentBar {
	readonly #deps: ExternalContentDeps;
	#consent = $state.raw<Consent>({ load: -1, enabled: false, closed: false });
	readonly #view: ExternalContentView = $derived.by(() => {
		const own = this.#consent.load === this.#deps.getLoadCount();
		return resolveExternalContentView({
			host: this.#deps.getHost(),
			options: this.#deps.getOptions(),
			enabled: own && this.#consent.enabled,
			closed: own && this.#consent.closed,
			content: this.#deps.getContent(),
		});
	});

	constructor(deps: ExternalContentDeps) {
		this.#deps = deps;
	}

	/** Surfaces draw the linked content (reactive). */
	get allowed(): boolean {
		return this.#view.allowed;
	}

	/** Whether the security bar shows (reactive). */
	get barVisible(): boolean {
		return this.#view.barVisible;
	}

	/** Enable Content: draw this presentation's linked content, for this viewer only. */
	readonly enable = (): void => {
		this.#consent = { ...this.#own(), enabled: true };
	};

	/** Close: hide the bar; the content stays blocked. */
	readonly close = (): void => {
		this.#consent = { ...this.#own(), closed: true };
	};

	#own(): Consent {
		const load = this.#deps.getLoadCount();
		return this.#consent.load === load ? this.#consent : { load, enabled: false, closed: false };
	}
}
