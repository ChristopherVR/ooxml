import { createPresentationRehearsalController } from 'ooxml-ui/pptx';
import type { PresentationRehearsalState } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import type { Store, ViewerState } from '../state';

export interface RehearseTimingsOptions {
	store: Store<ViewerState>;
	/** Open the slide show the rehearsal records. */
	enterPresentation(): Promise<void>;
	onSave(timings: Record<number, number>): void;
}

const format = (ms: number): string => new Date(Math.max(0, ms)).toISOString().slice(14, 19);

/**
 * Rehearse Timings for the vanilla binding: runs the slide show with the
 * shared `createPresentationRehearsalController` timing it (the controller owns
 * the clocks, pause and the display tick), and paints its HUD and the
 * keep-the-timings summary. Returns the function that tears it down.
 */
export function openRehearseTimings(
	doc: Document,
	host: HTMLElement,
	t: Translator,
	options: RehearseTimingsOptions,
): () => void {
	host.querySelector('.pptxv-rehearse')?.remove();
	const rehearsal = createPresentationRehearsalController();
	const panel = createEl(doc, 'aside', 'pptxv-rehearse');
	panel.setAttribute('aria-label', t('pptx.rehearse.summaryTitle'));
	host.appendChild(panel);

	const button = (label: string, action: () => void): HTMLButtonElement => {
		const el = createEl(doc, 'button');
		el.type = 'button';
		el.textContent = label;
		el.addEventListener('click', action);
		return el;
	};
	const text = (tag: 'span' | 'strong', value: string): HTMLElement => {
		const el = createEl(doc, tag);
		el.textContent = value;
		return el;
	};
	let teardown = (): void => undefined;
	const close = (): void => teardown();

	const paint = (state: PresentationRehearsalState): void => {
		if (state.active) {
			panel.setAttribute('role', 'status');
			panel.replaceChildren(
				text('span', `${t('pptx.rehearse.slide')} ${state.slideIndex + 1}`),
				text('strong', `${t('pptx.rehearse.slideTime')}: ${format(state.slideElapsedMs)}`),
				text('span', `${t('pptx.rehearse.totalTime')}: ${format(state.totalElapsedMs)}`),
				button(t(state.paused ? 'pptx.rehearse.resume' : 'pptx.rehearse.pause'), () =>
					rehearsal.togglePause(),
				),
			);
			return;
		}
		if (state.summaryOpen) {
			panel.setAttribute('role', 'dialog');
			const rows = Object.entries(state.timings).map(([index, ms]) =>
				text('span', `${t('pptx.rehearse.slide')} ${Number(index) + 1}: ${format(ms)}`),
			);
			panel.replaceChildren(
				...rows,
				text(
					'strong',
					`${t('pptx.rehearse.totalPresentationTime')}: ${format(state.recordedTotalMs)}`,
				),
				button(t('pptx.rehearse.saveTimings'), () => {
					options.onSave({ ...rehearsal.save() });
					close();
				}),
				button(t('pptx.rehearse.discard'), () => {
					rehearsal.discard();
					close();
				}),
			);
		}
	};

	const unsubscribeRehearsal = rehearsal.subscribe(() => paint(rehearsal.getSnapshot()));
	const unsubscribeStore = options.store.subscribe((state, previous) => {
		if (state.currentSlide !== previous.currentSlide) {
			rehearsal.moveTo(state.currentSlide);
		}
		if (previous.presenting && !state.presenting) {
			rehearsal.finish();
		}
	});
	teardown = () => {
		unsubscribeRehearsal();
		unsubscribeStore();
		rehearsal.dispose();
		panel.remove();
	};
	// Start the clock once the show is up, on the slide it actually opened on.
	void options.enterPresentation().then(() => {
		rehearsal.start(options.store.get().currentSlide);
		return undefined;
	});
	return close;
}
