/**
 * `usePresenterSession`: the presenter and audience windows for the Vue
 * binding. The window, deck hand-over, channel protocol and Swap Displays are
 * the shared `createPresenterSessionController`; this composable ties it to
 * the component lifecycle and the current slide.
 *
 * @module composables/usePresenterSession
 */

import { createPresenterSessionController } from 'ooxml-ui/pptx';
import type { PresentationSnapshot } from 'ooxml-ui/pptx';
import { computed, onBeforeUnmount, onMounted, watch } from 'vue';
import type { Ref } from 'vue';

import { useController } from './useController';

export interface PresenterSessionOptions {
	currentSlideIndex: Ref<number>;
	content: () => ArrayBuffer | Uint8Array | null;
	onAudienceSlide: (index: number) => void;
	onAudienceExit: () => void;
}

export function usePresenterSession(options: PresenterSessionOptions) {
	const session = createPresenterSessionController({
		getSource: options.content,
		initialSlideIndex: options.currentSlideIndex.value,
		onAudienceSlide: options.onAudienceSlide,
		onAudienceExit: options.onAudienceExit,
	});
	const state = useController(session);
	const isAudience = session.getSnapshot().role === 'audience';

	watch(options.currentSlideIndex, (slideIndex) => {
		if (!isAudience) {
			session.updateSnapshot({ slideIndex });
		}
	});
	onMounted(session.connect);
	onBeforeUnmount(session.dispose);

	return {
		isAudience,
		audienceOpen: computed(() => state.value.audienceOpen),
		snapshot: computed<PresentationSnapshot>(() => state.value.snapshot),
		updateSnapshot: session.updateSnapshot,
		openAudience: session.openAudience,
		closeAudience: session.closeAudience,
		swapDisplays: session.swapDisplays,
	};
}
