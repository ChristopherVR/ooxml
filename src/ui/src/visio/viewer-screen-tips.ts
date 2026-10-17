/** Office's ScreenTip style: full tips, the name alone, or none. */
export type ScreenTipStyle = 'descriptions' | 'plain' | 'off';
export const isScreenTipStyle = (value: unknown): value is ScreenTipStyle =>
	value === 'descriptions' || value === 'plain' || value === 'off';

const FULL = 'data-screen-tip';
/** A command that is not available explains why in its tip; that explanation is never hidden. */
const isReason = (tip: string) => tip.includes(': not available yet.');

/** Put back every tip `applyScreenTips` set aside (before the ribbon is translated again). */
export function restoreScreenTips(root: ParentNode): void {
	for (const element of root.querySelectorAll<HTMLElement>(`[${FULL}]`)) {
		element.setAttribute('title', element.getAttribute(FULL)!);
		element.removeAttribute(FULL);
	}
}

/**
 * Applies the ScreenTip style to a ribbon's `title` tooltips: "plain" drops the shortcut hint,
 * "off" removes the tip. The original is kept so the style can change back.
 */
export function applyScreenTips(root: ParentNode, style: ScreenTipStyle): void {
	restoreScreenTips(root);
	if (style === 'descriptions') return;
	for (const element of root.querySelectorAll<HTMLElement>('[title]')) {
		const tip = element.getAttribute('title')!;
		if (isReason(tip)) continue;
		element.setAttribute(FULL, tip);
		if (style === 'off') element.removeAttribute('title');
		else element.setAttribute('title', tip.replace(/ \([^()]*\)$/, ''));
	}
}
