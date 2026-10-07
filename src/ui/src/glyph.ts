// An inline icon template for Lit elements: the registry glyph drawn as an `<svg>`. An unknown
// name renders an empty, unpainted svg (`data-painted` absent) so a stylesheet can hide it. Lit
// only rewrites the `d` attribute when the glyph changes, so a repaint under a pressed pointer
// never replaces the node (which would make the browser drop the click). The svg holds no
// whitespace, so it never adds text to its parent's `textContent`.
import { html, svg, type TemplateResult } from 'lit';
import { getIcon } from './icons';

export function glyph(
	name: string | null | undefined,
	className = 'glyph',
	hidden = false,
): TemplateResult {
	const found = getIcon(name);
	// The formatter would indent the path onto its own line; the svg must hold no whitespace text.
	// prettier-ignore
	return html`<svg
		class=${className}
		viewBox=${found?.viewBox ?? '0 0 20 20'}
		aria-hidden="true"
		?data-painted=${Boolean(found)}
		?hidden=${hidden}
	>${found ? svg`<path d=${found.d}></path>` : ''}</svg>`;
}
