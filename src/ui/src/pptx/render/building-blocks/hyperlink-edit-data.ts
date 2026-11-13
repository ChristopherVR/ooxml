/**
 * The editor values of a hyperlink (the Insert Hyperlink dialog's fields) and their conversion to
 * and from the `actionClick` an element stores. Every binding's dialog and the building blocks'
 * `actions.hyperlink` use the same shape, so a host editor written against one binding works
 * against all of them. DOM-free and framework-free.
 *
 * @module render/building-blocks/hyperlink-edit-data
 */

/** The type of hyperlink target. */
export type HyperlinkTargetType = 'url' | 'email' | 'slide' | 'file' | 'action';

/** Supported action verbs for action hyperlinks. */
export type HyperlinkActionVerb =
	| 'nextSlide'
	| 'prevSlide'
	| 'firstSlide'
	| 'lastSlide'
	| 'endShow';

/** Data model for the hyperlink editor dialog. */
export interface HyperlinkEditData {
	targetType: HyperlinkTargetType;
	url: string;
	tooltip: string;
	emailAddress: string;
	emailSubject: string;
	slideNumber: number;
	filePath: string;
	actionVerb: HyperlinkActionVerb;
}

export const HYPERLINK_ACTION_VERB_MAP: Record<string, HyperlinkActionVerb> = {
	'ppaction://hlinkshowjump?jump=nextslide': 'nextSlide',
	'ppaction://hlinkshowjump?jump=previousslide': 'prevSlide',
	'ppaction://hlinkshowjump?jump=firstslide': 'firstSlide',
	'ppaction://hlinkshowjump?jump=lastslide': 'lastSlide',
	'ppaction://hlinkshowjump?jump=endshow': 'endShow',
};

export const HYPERLINK_ACTION_VERB_TO_PPACTION: Record<HyperlinkActionVerb, string> = {
	nextSlide: 'ppaction://hlinkshowjump?jump=nextslide',
	prevSlide: 'ppaction://hlinkshowjump?jump=previousslide',
	firstSlide: 'ppaction://hlinkshowjump?jump=firstslide',
	lastSlide: 'ppaction://hlinkshowjump?jump=lastslide',
	endShow: 'ppaction://hlinkshowjump?jump=endshow',
};

export function detectHyperlinkTargetType(
	url: string | undefined,
	action: string | undefined,
): HyperlinkTargetType {
	if (action && action.startsWith('ppaction://')) {
		if (action.includes('hlinksldjump')) {
			return 'slide';
		}
		return 'action';
	}
	if (!url) {
		return 'url';
	}
	if (url.startsWith('mailto:')) {
		return 'email';
	}
	if (/^[a-zA-Z]:\\|^\.\.?[/\\]|^file:/i.test(url)) {
		return 'file';
	}
	return 'url';
}

export function parseHyperlinkEmailUrl(url: string): { address: string; subject: string } {
	if (!url.startsWith('mailto:')) {
		return { address: url, subject: '' };
	}
	const withoutScheme = url.slice(7);
	const qIdx = withoutScheme.indexOf('?');
	if (qIdx < 0) {
		return { address: withoutScheme, subject: '' };
	}
	const address = withoutScheme.slice(0, qIdx);
	const params = new URLSearchParams(withoutScheme.slice(qIdx + 1));
	return { address, subject: params.get('subject') || '' };
}

export function parseHyperlinkSlideNumber(
	url: string | undefined,
	action: string | undefined,
): number {
	if (action === 'ppaction://hlinksldjump' && url) {
		const match = url.match(/slide(\d+)\.xml$/i);
		if (match?.[1]) {
			return parseInt(match[1], 10);
		}
	}
	if (url) {
		const hashMatch = url.match(/#\s*Slide\s+(\d+)/i);
		if (hashMatch?.[1]) {
			return parseInt(hashMatch[1], 10);
		}
	}
	return 1;
}

/** Resolve hyperlink edit data to a URL string and optional action for the element. */
export function resolveHyperlinkEditResult(data: HyperlinkEditData): {
	url: string;
	action?: string;
	tooltip?: string;
} {
	const tooltip = data.tooltip.trim() || undefined;
	const withTooltip = <T extends { url: string; action?: string }>(value: T) =>
		tooltip === undefined ? value : { ...value, tooltip };
	switch (data.targetType) {
		case 'email': {
			const mailto = data.emailSubject
				? `mailto:${data.emailAddress}?subject=${encodeURIComponent(data.emailSubject)}`
				: `mailto:${data.emailAddress}`;
			return withTooltip({ url: mailto });
		}
		case 'slide':
			return withTooltip({
				url: `slide${data.slideNumber}.xml`,
				action: 'ppaction://hlinksldjump',
			});
		case 'file':
			return withTooltip({ url: data.filePath });
		case 'action':
			return withTooltip({ url: '', action: HYPERLINK_ACTION_VERB_TO_PPACTION[data.actionVerb] });
		case 'url':
		default:
			return withTooltip({ url: data.url });
	}
}

/**
 * The editor values for an element's current link, detected the way the hyperlink dialog seeds
 * its fields from it. A host-rendered editor starts from this and passes the edited value back.
 */
export function hyperlinkEditDataFromAction(
	actionClick: { url?: string; action?: string; tooltip?: string } | undefined,
): HyperlinkEditData | null {
	if (!actionClick || (!actionClick.url && !actionClick.action)) {
		return null;
	}
	const { url, action, tooltip } = actionClick;
	const targetType = detectHyperlinkTargetType(url, action);
	const email = parseHyperlinkEmailUrl(url ?? '');
	return {
		targetType,
		url: targetType === 'url' ? (url ?? '') : '',
		tooltip: tooltip ?? '',
		emailAddress: targetType === 'email' ? email.address : '',
		emailSubject: targetType === 'email' ? email.subject : '',
		slideNumber: parseHyperlinkSlideNumber(url, action),
		filePath: targetType === 'file' ? (url ?? '') : '',
		actionVerb: (action && HYPERLINK_ACTION_VERB_MAP[action.toLowerCase()]) || 'nextSlide',
	};
}
