import {
	applyAnimationPaint,
	captureAnimationPaint,
	type AnimationPaint,
} from '../animation-painter';
import type { RibbonAnimationEditing } from '../ribbon-animation-editing';
import { createControllerStore } from './controller';

type PainterState = { mode: 'off' } | { mode: 'once' | 'locked'; paint: AnimationPaint };

export function createAnimationPainterController() {
	const off: PainterState = { mode: 'off' };
	const store = createControllerStore<PainterState>(off);
	const cancel = (): void => store.set(off);
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		cancel,
		arm(editing: RibbonAnimationEditing | undefined, mode: 'once' | 'locked'): void {
			const current = store.getSnapshot();
			if (current.mode !== 'off') {
				if (mode === 'locked' && current.mode === 'once') {
					store.set({ mode, paint: current.paint });
				}
				return;
			}
			if (editing?.commit && editing.slide) {
				const paint = captureAnimationPaint(editing.slide, editing.elementId);
				if (paint) {
					store.set({ mode, paint });
				}
			}
		},
		apply(editing: RibbonAnimationEditing | undefined, elementId: string): void {
			const current = store.getSnapshot();
			if (current.mode === 'off' || !editing?.commit || !editing.slide) {
				return;
			}
			const next = applyAnimationPaint(current.paint, editing.slide, elementId);
			if (next) {
				if (current.mode === 'once') {
					cancel();
				}
				editing.commit(next);
			}
		},
	};
}
