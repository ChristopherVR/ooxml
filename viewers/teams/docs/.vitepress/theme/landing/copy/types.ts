import type { IconName } from '../icons';

export interface LandingLink {
	text: string;
	/** Site-relative path (without the base) or an absolute URL. */
	href: string;
}

export interface LandingCardCopy {
	title: string;
	copy: string;
	href: string;
	link: string;
	icon?: IconName;
}

export interface LandingFaqItem {
	q: string;
	a: string;
	href?: string;
	link?: string;
}

export interface LandingFooterColumn {
	title: string;
	links: Array<LandingLink & { external: boolean }>;
}

/** Every user-visible string of the Teams landing page, one dictionary per locale. */
export interface LandingCopy {
	hero: {
		titleTop: string;
		titleAccent: string;
		sub: string;
		start: LandingLink;
		demo: LandingLink;
		frameLabel: string;
		tryLabel: string;
		imageAlt: string;
	};
	liveDemo: {
		title: string;
		note: string;
		introBefore: string;
		introLink: LandingLink;
		introAfter: string;
		/** aria-label of the framework tablist. */
		frameworkLabel: string;
		soloTab: string;
		duoTab: string;
		guestPicker: string;
		load: string;
		openFull: string;
		/** Title of the embedded frame; `{framework}` is replaced. */
		frameTitle: string;
		hintDuo: string;
		hintSolo: string;
	};
	embed: {
		title: string;
		note: string;
		openDemo: string;
		demos: LandingCardCopy[];
	};
	features: {
		title: string;
		note: string;
		items: LandingCardCopy[];
	};
	limits: {
		title: string;
		note: string;
		items: string[];
		all: LandingLink;
	};
	quickstart: {
		title: string;
		note: string;
		/** Text before and after the `config` code word. */
		introBefore: string;
		introAfter: string;
		/** `{framework}` is replaced. */
		fullGuide: string;
	};
	faq: {
		title: string;
		note: string;
		items: LandingFaqItem[];
	};
	finale: {
		title: string;
		introBefore: string;
		introAfter: string;
		start: LandingLink;
		github: string;
		columns: LandingFooterColumn[];
		license: string;
		disclaimer: string;
	};
}
