type CueModule = Pick<typeof import('./ribbon-tab-cue'), 'wireRibbonTabCue'>;

/**
 * Loads the tab-strip overflow cue (`ribbon-tab-cue.ts`) on demand for a ribbon with an
 * `overflow-label`. A failed chunk load is not final: the next `ensure()`, which the ribbon calls
 * on every update, tries again.
 */
export class RibbonTabCueLoader {
	private controller: AbortController | undefined;
	private readonly ribbon: HTMLElement;
	private readonly load: () => Promise<CueModule>;

	constructor(
		ribbon: HTMLElement,
		load: () => Promise<CueModule> = () => import('./ribbon-tab-cue'),
	) {
		this.ribbon = ribbon;
		this.load = load;
	}

	ensure(): void {
		if (this.controller || !this.ribbon.hasAttribute('overflow-label')) return;
		const controller = (this.controller = new AbortController());
		const { signal } = controller;
		void this.load().then(
			({ wireRibbonTabCue }) => {
				if (!signal.aborted)
					signal.addEventListener('abort', wireRibbonTabCue(this.ribbon), { once: true });
			},
			() => {
				if (this.controller === controller) this.controller = undefined;
			},
		);
	}

	stop(): void {
		this.controller?.abort();
		this.controller = undefined;
	}
}
