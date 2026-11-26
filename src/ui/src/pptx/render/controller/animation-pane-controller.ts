import {
	getAnimationPresetInfo,
	hasTextProperties,
	nativeAnimationClickOrders,
	type PptxElement,
} from 'ooxml-core/pptx';

import { animationCatalogPresetLabelKey } from '../animation-preset-labels';
import type { AnimationEditingSlide } from '../ribbon-animation-editing';
import { animationsLabel, type RibbonAnimationsViewState } from '../ribbon-animations-state';
import { createControllerStore } from './controller';

interface AnimationPaneRow {
	nodeId: number;
	clickOrder: number | undefined;
	label: string;
}

interface AnimationPaneSnapshot {
	nodeId?: number | undefined;
	context?: RibbonAnimationsViewState | undefined;
	nativeAnimations?: AnimationEditingSlide['nativeAnimations'];
	rows: readonly AnimationPaneRow[];
}

function shapeLabels(elements: readonly PptxElement[], labels = new Map<string, string>()) {
	for (const element of elements) {
		const text = hasTextProperties(element) ? element.text?.trim() : undefined;
		labels.set(element.id, element.name?.trim() || text || element.type);
		if (element.type === 'group') shapeLabels(element.children, labels);
	}
	return labels;
}

function buildNativeRows(context: RibbonAnimationsViewState): AnimationPaneRow[] {
	const slide = context.editing?.slide;
	if (!slide?.rawTiming) return [];
	const orders = nativeAnimationClickOrders(slide.rawTiming);
	const anchors = new Set(slide.animationTimelineAnchors?.map((anchor) => anchor.order));
	const labels = shapeLabels(slide.elements ?? []);
	return (slide.nativeAnimations ?? []).flatMap((effect) => {
		const { nodeId, targetId, presetClass, presetId } = effect;
		if (nodeId === undefined || !targetId || !presetClass || !orders.has(nodeId)) return [];
		const order = orders.get(nodeId);
		const token = `${presetClass}.${presetId}`;
		const preset = getAnimationPresetInfo(token);
		const effectLabel = animationsLabel(
			context,
			animationCatalogPresetLabelKey(token),
			preset?.label ?? animationsLabel(context, 'pptx.animation.nativeEffect', 'PowerPoint effect'),
		);
		return [
			{
				nodeId,
				clickOrder: order !== undefined && anchors.has(order) ? order : undefined,
				label: `${labels.get(targetId) ?? targetId} · ${effectLabel}`,
			},
		];
	});
}

/** Selection shared by the Animation Pane and the Animations ribbon. */
export function createAnimationPaneController() {
	const store = createControllerStore<AnimationPaneSnapshot>({ rows: [] });
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		sync(context: RibbonAnimationsViewState) {
			const previous = store.getSnapshot();
			const nativeAnimations = context.editing?.slide?.nativeAnimations;
			const sameEffects = previous.nativeAnimations === nativeAnimations;
			if (previous.context === context && sameEffects) return;
			const nodeId =
				context.paneOpen &&
				sameEffects &&
				context.editing?.slide?.id === previous.context?.editing?.slide?.id &&
				context.editing?.elementId === previous.context?.editing?.elementId
					? previous.nodeId
					: undefined;
			const rows =
				sameEffects &&
				previous.context?.editing?.slide === context.editing?.slide &&
				previous.context?.translate === context.translate
					? previous.rows
					: buildNativeRows(context);
			store.set({ nodeId, context, nativeAnimations, rows });
		},
		select(nodeId: number | undefined) {
			const previous = store.getSnapshot();
			if (nodeId === previous.nodeId) return;
			if (nodeId !== undefined && !previous.rows.some((row) => row.nodeId === nodeId)) return;
			store.set({ ...previous, nodeId });
		},
	};
}
