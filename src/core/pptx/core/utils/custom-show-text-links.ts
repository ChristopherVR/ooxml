import type { TextStyle } from '../types';
import { rewriteCustomShowAction, type CustomShowActionRewrite } from './custom-show-link';

const CLICK_KEYS = [
	'hyperlink',
	'hyperlinkRId',
	'hyperlinkTooltip',
	'hyperlinkAction',
	'hyperlinkTargetSlideIndex',
	'hyperlinkSoundXml',
	'hyperlinkInvalidUrl',
	'hyperlinkTargetFrame',
	'hyperlinkHistory',
	'hyperlinkHighlightClick',
	'hyperlinkEndSound',
] as const satisfies readonly (keyof TextStyle)[];

/** Rewrites canonical text links and their action metadata without changing parse provenance. */
export function mapCustomShowTextLinks(style: TextStyle, rewrite: CustomShowActionRewrite) {
	const url = rewriteCustomShowAction(style.hyperlink, rewrite);
	const action = rewriteCustomShowAction(style.hyperlinkAction, rewrite);
	const hover = rewriteCustomShowAction(style.hyperlinkMouseOver, rewrite);
	const clickChanged = url.changed || action.changed;
	if (!clickChanged && !hover.changed) return { style, click: false, hover: false };
	const next = { ...style };
	if (clickChanged) {
		if (
			(url.changed && url.action === undefined) ||
			(action.changed && action.action === undefined)
		) {
			for (const key of CLICK_KEYS) delete next[key];
		} else {
			if (url.changed && url.action !== undefined) next.hyperlink = url.action;
			if (action.changed && action.action !== undefined) {
				next.hyperlinkAction = action.action;
				if (!style.hyperlink) next.hyperlink = action.action;
			}
		}
	}
	if (hover.changed) {
		if (hover.action === undefined) {
			delete next.hyperlinkMouseOver;
			delete next.hyperlinkMouseOverSoundXml;
		} else {
			next.hyperlinkMouseOver = hover.action;
		}
	}
	return { style: next, click: clickChanged, hover: hover.changed };
}
