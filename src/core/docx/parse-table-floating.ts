// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads `w:tblpPr` (ECMA-376 17.4.58), the position of a floating table.
import type { TableFloatingPosition } from './table-model';
import { getW, type XmlElement } from './xml';
import { parseSignedTwips, parseTwips } from './simple-types';

const ANCHORS = ['margin', 'page', 'text'] as const;
const X_ALIGN = ['left', 'center', 'right', 'inside', 'outside'] as const;
const Y_ALIGN = ['top', 'center', 'bottom', 'inside', 'outside', 'inline'] as const;

function oneOf<T extends string>(values: readonly T[], value: string | undefined): T | undefined {
	return values.find((item) => item === value);
}

/** The floating position of a table, or undefined when it has no `w:tblpPr`. */
export function parseTableFloating(
	element: XmlElement | undefined,
): TableFloatingPosition | undefined {
	if (!element) return undefined;
	const position: TableFloatingPosition = {};
	const horizontalAnchor = oneOf(ANCHORS, getW(element, 'horzAnchor'));
	if (horizontalAnchor) position.horizontalAnchor = horizontalAnchor;
	const verticalAnchor = oneOf(ANCHORS, getW(element, 'vertAnchor'));
	if (verticalAnchor) position.verticalAnchor = verticalAnchor;
	const x = parseSignedTwips(getW(element, 'tblpX'));
	if (x !== undefined) position.xTwips = x;
	const y = parseSignedTwips(getW(element, 'tblpY'));
	if (y !== undefined) position.yTwips = y;
	const xAlign = oneOf(X_ALIGN, getW(element, 'tblpXSpec'));
	if (xAlign) position.xAlign = xAlign;
	const yAlign = oneOf(Y_ALIGN, getW(element, 'tblpYSpec'));
	if (yAlign) position.yAlign = yAlign;
	for (const [attribute, key] of [
		['leftFromText', 'leftFromTextTwips'],
		['rightFromText', 'rightFromTextTwips'],
		['topFromText', 'topFromTextTwips'],
		['bottomFromText', 'bottomFromTextTwips'],
	] as const) {
		const value = parseTwips(getW(element, attribute));
		if (value !== undefined) position[key] = value;
	}
	return position;
}
