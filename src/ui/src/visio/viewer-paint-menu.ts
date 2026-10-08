import { visioShapeFormattingState } from 'ooxml-core/visio/ui';
import type { VisioShape } from 'ooxml-core/visio';
import type { CommandSpec, RibbonCommand } from './ribbon-parts';
import { paintOptions } from './ribbon-style-options';

/** Paint menu checks use source property state; mixed selections have no checked choice. */
export function renderPaintMenus(
	root: ShadowRoot,
	shapes: readonly VisioShape[],
	reason: string,
	set: (element: RibbonCommand | null, reason: string) => void,
): void {
	const common = visioShapeFormattingState(shapes);
	const visit = (spec: CommandSpec, target: 'fill' | 'line') => {
		if (spec.items) {
			set(root.querySelector(`[data-menu="${spec.id}"]`), reason);
			for (const item of spec.items) visit(item, target);
			return;
		}
		const element = root.querySelector<RibbonCommand>(`[command="${spec.id}"]`);
		set(element, reason);
		if (!element || spec.action?.type !== 'shape-format') return;
		const patch = spec.action.patch;
		const checked =
			!reason &&
			shapes.length > 0 &&
			(patch.linePattern !== undefined
				? common.linePattern === patch.linePattern
				: patch.lineWeight !== undefined
					? shapes.every(
							(shape) => Math.abs(shape.style.lineWidth * 72 - patch.lineWeight!) < 0.001,
						)
					: shapes.every(
							(shape) =>
								(target === 'fill' ? shape.style.fill : shape.style.lineColor) ===
								(target === 'fill' ? patch.fillColor : patch.lineColor),
						));
		element.setAttribute('checked', String(checked));
	};
	for (const target of ['fill', 'line'] as const) {
		set(root.querySelector(`[data-menu="${target}"]`), reason);
		for (const spec of paintOptions(target)) visit(spec, target);
	}
}
