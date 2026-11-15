import { twipsToPixels, type SignedTwips } from '../index';
import type { TableFloatingPosition } from '../table-model';
import type { LayoutTableFloating } from './input';

/**
 * A floating table's frame for layout. Measured in Word 16 (floating-table fixtures): `text` is the
 * text column horizontally and, vertically, the top of the paragraph that follows the table; an
 * absent `w:vertAnchor` measures from the top margin, and an absent `w:horzAnchor` from the column
 * (the same as the margin for a single-column page). `tblpXSpec`/`tblpYSpec` win over the offsets;
 * `inline` keeps the table at its place in the flow.
 */
export function floatingFrame(position: TableFloatingPosition): LayoutTableFloating {
	const px = (twips: SignedTwips | undefined) => (twips === undefined ? 0 : twipsToPixels(twips));
	const relativeFromH =
		position.horizontalAnchor === 'page' || position.horizontalAnchor === 'margin'
			? position.horizontalAnchor
			: 'column';
	const relativeFromV =
		position.verticalAnchor === 'page'
			? 'page'
			: position.verticalAnchor === 'text' || position.yAlign === 'inline'
				? 'paragraph'
				: 'margin';
	const alignV = position.yAlign === 'inline' ? undefined : position.yAlign;
	return {
		relativeFromH,
		...(position.xAlign ? { alignH: position.xAlign } : {}),
		offsetXPx: px(position.xTwips),
		relativeFromV,
		...(alignV ? { alignV } : {}),
		offsetYPx: position.yAlign === 'inline' ? 0 : px(position.yTwips),
		distance: {
			topPx: px(position.topFromTextTwips),
			rightPx: px(position.rightFromTextTwips),
			bottomPx: px(position.bottomFromTextTwips),
			leftPx: px(position.leftFromTextTwips),
		},
	};
}
