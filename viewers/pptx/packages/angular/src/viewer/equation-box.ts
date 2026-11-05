import { buildEquationBlockStyle, isEquationOnlyText } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

import { getTextBlockStyle } from './element-style';
import type { StyleMap } from './element-style';

/** A pure equation box: the shared block style plus the equations to paint. */
export interface EquationBox {
	readonly style: StyleMap;
	readonly equations: ReadonlyArray<{
		readonly xml: Record<string, unknown>;
		readonly number: string | undefined;
	}>;
}

/**
 * Resolve a pure equation box (OMML and nothing else) through the shared
 * `buildEquationBlockStyle` descriptor, the one Vue, Svelte and Vanilla paint.
 * Returns `undefined` for any other body (prose with a formula stays on the
 * paragraph path, where the equation is a run in place).
 */
export function resolveEquationBox(el: PptxElement): EquationBox | undefined {
	if (!hasTextProperties(el) || !isEquationOnlyText(el.textSegments)) {
		return undefined;
	}
	const equations = (el.textSegments ?? []).flatMap((segment) =>
		segment.equationXml
			? [{ xml: segment.equationXml as Record<string, unknown>, number: segment.equationNumber }]
			: [],
	);
	if (equations.length === 0) {
		return undefined;
	}
	return { style: buildEquationBlockStyle(getTextBlockStyle(el)), equations };
}
