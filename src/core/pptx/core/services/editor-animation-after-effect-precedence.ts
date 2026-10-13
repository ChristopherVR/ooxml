import type { PptxElementAnimation, PptxNativeAnimation, XmlObject } from '../types';
import { assertAfterAnimationClearSupported } from './animation-after-effect-write';
import { hasAuthoredAnimationEffect } from './editor-animation-entry';

export function applyAfterAnimationFromEditorList(
	nativeAnimations: readonly PptxNativeAnimation[],
	editorAnimations: readonly PptxElementAnimation[] | undefined,
): PptxNativeAnimation[] {
	if (!editorAnimations || editorAnimations.length === 0) {
		return [...nativeAnimations];
	}
	const byElement = new Map<string, PptxElementAnimation>();
	for (const ea of editorAnimations) {
		if (ea.afterAnimation !== undefined) {
			byElement.set(ea.elementId, ea);
		}
	}
	if (byElement.size === 0) {
		return [...nativeAnimations];
	}
	for (const [id, editor] of byElement) {
		const matches = nativeAnimations.filter(
			(native) => native.targetId === id && ['entr', 'emph'].includes(native.presetClass ?? ''),
		);
		if (matches.length > 1)
			throw new Error(
				'Cannot edit an ambiguous native after-animation target with multiple effects',
			);
		if (matches.length === 0 && !hasAuthoredAnimationEffect(editor))
			throw new Error('Cannot edit after-animation without a supported native effect');
		for (const native of matches) {
			if (
				native.textTarget ||
				(native.target &&
					(native.target.type !== 'shape' ||
						native.target.backgroundOnly ||
						native.target.subShapeId ||
						native.target.graphicElement ||
						native.target.oleChartElement))
			)
				throw new Error('Editing after-animation supports whole-shape targets only');
			assertAfterAnimationClearSupported({
				'p:subTnLst': native.cTnAttributes?.['p:subTnLst'],
			} as XmlObject);
		}
	}
	return nativeAnimations.map((anim) => {
		if (!anim.targetId || !['entr', 'emph'].includes(anim.presetClass ?? '')) {
			return anim;
		}
		const ea = byElement.get(anim.targetId);
		if (!ea) {
			return anim;
		}
		return {
			...anim,
			afterAnimationAction: ea.afterAnimation === 'none' ? undefined : ea.afterAnimation,
			afterAnimationColor: ea.afterAnimationColor,
			// The editor's own model has no theme-colour-ref concept (it always
			// writes a concrete hex); clear a native scheme ref so a stale one
			// never survives an edit that overrides the colour.
			afterAnimationColorRef: undefined,
		};
	});
}
