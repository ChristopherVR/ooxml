import type { PageSetup } from '../model';
import type { HeaderFooterSection } from './header-footer';

export type HeaderFooterKind = 'header' | 'footer';
export type HeaderFooterKey =
	| 'header'
	| 'footer'
	| 'evenHeader'
	| 'evenFooter'
	| 'firstHeader'
	| 'firstFooter';

/**
 * Splits a header or footer string into the raw text of its left, centre and right sections
 * (formatting and field codes such as `&P` kept), the inverse of `composeHeaderFooter`.
 */
export function splitHeaderFooter(value: string): Record<HeaderFooterSection, string> {
	const result: Record<HeaderFooterSection, string> = { left: '', center: '', right: '' };
	let section: HeaderFooterSection = 'center';
	for (let i = 0; i < value.length;) {
		const ch = value[i]!;
		if (ch !== '&') {
			result[section] += ch;
			i++;
			continue;
		}
		const code = value[i + 1];
		if (code === 'L' || code === 'C' || code === 'R') {
			section = code === 'L' ? 'left' : code === 'C' ? 'center' : 'right';
			i += 2;
		} else if (code === '"') {
			const end = value.indexOf('"', i + 2);
			const stop = end < 0 ? value.length : end + 1;
			result[section] += value.slice(i, stop);
			i = stop;
		} else {
			result[section] += value.slice(i, i + 2);
			i += 2;
		}
	}
	return result;
}

/** Joins section texts into `&L...&C...&R...`; empty sections are left out. */
export function composeHeaderFooter(
	sections: Partial<Record<HeaderFooterSection, string>>,
): string {
	return (
		[
			['L', sections.left],
			['C', sections.center],
			['R', sections.right],
		] as const
	)
		.filter(([, text]) => text)
		.map(([code, text]) => `&${code}${text}`)
		.join('');
}

/** The page-setup field that supplies the header or footer of a page. */
export function headerFooterKey(
	setup: PageSetup,
	kind: HeaderFooterKind,
	physicalPage: number,
	pageNumber: number,
): HeaderFooterKey {
	const head = kind === 'header';
	if (physicalPage === 1 && setup.differentFirst) return head ? 'firstHeader' : 'firstFooter';
	if (pageNumber % 2 === 0 && setup.differentOddEven) return head ? 'evenHeader' : 'evenFooter';
	return kind;
}

/** A page-setup patch that replaces one section of a header or footer. */
export function patchHeaderFooterSection(
	setup: PageSetup,
	key: HeaderFooterKey,
	section: HeaderFooterSection,
	text: string,
): Partial<PageSetup> {
	const sections = splitHeaderFooter(setup[key] ?? '');
	sections[section] = text;
	const composed = composeHeaderFooter(sections);
	return { [key]: composed || undefined };
}
