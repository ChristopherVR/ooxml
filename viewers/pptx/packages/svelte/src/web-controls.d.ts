import type { HTMLAttributes, HTMLInputAttributes, HTMLSelectAttributes } from 'svelte/elements';

declare module 'svelte/elements' {
	interface SVGAttributes<T> {
		slot?: string;
	}
	interface SvelteHTMLElements {
		'pptx-ui-select': HTMLSelectAttributes & { variant?: 'ribbon-font' | 'ribbon-icon' };
		'pptx-ui-checkbox': HTMLInputAttributes;
		'pptx-ui-animation-effect-rows': HTMLAttributes<HTMLElement> & {
			'group-order'?: number;
			all?: string;
		};
	}
}
