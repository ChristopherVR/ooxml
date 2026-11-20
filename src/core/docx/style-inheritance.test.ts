import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './parse';
import { chainResolver, inheritanceCycleEntries } from './style-inheritance';
import { docx, p } from './test-support/compare-docs';
import { growthAsync } from '../test-support/growth';

interface Style {
	basedOn?: string;
	value: Record<string, number>;
}

/** The walk every resolver replaced: nearest first, stopping at a missing style or a repeat. */
function naiveChain(styles: Record<string, Style>, id: string): string[] {
	const chain: string[] = [];
	let current: string | undefined = id;
	while (current && styles[current] && !chain.includes(current)) {
		chain.push(current);
		current = styles[current]!.basedOn;
	}
	return chain;
}

function naiveCycleEntries(styles: Record<string, Style>): string[] {
	const warnings: string[] = [];
	for (const id of Object.keys(styles)) {
		const seen = new Set([id]);
		for (
			let parent = styles[id]!.basedOn;
			parent && styles[parent];
			parent = styles[parent]!.basedOn
		) {
			if (seen.has(parent)) {
				warnings.push(parent);
				break;
			}
			seen.add(parent);
		}
	}
	return [...new Set(warnings)];
}

/** A random catalog with chains, shared ancestors, cycles and missing parents. */
function catalog(seed: number, size: number): Record<string, Style> {
	let state = seed;
	const next = (limit: number) => {
		state = (state * 1103515245 + 12345) & 0x7fffffff;
		return state % limit;
	};
	const styles: Record<string, Style> = Object.create(null);
	for (let index = 0; index < size; index++) {
		const roll = next(10);
		styles[`s${index}`] = {
			...(roll < 8 ? { basedOn: `s${next(roll < 6 ? size : size + 3)}` } : {}),
			value: { [`k${next(4)}`]: index, [`k${next(4)}`]: -index },
		};
	}
	return styles;
}

describe('style inheritance', () => {
	it('resolves every style as the chain walk it replaced does, cycles included', () => {
		for (let seed = 1; seed <= 60; seed++) {
			const styles = catalog(seed, 1 + (seed % 25));
			const resolve = chainResolver(
				styles,
				(style) => style.value,
				(resolved, value) => ({ ...resolved, ...value }),
				{} as Record<string, number>,
			);
			for (const id of Object.keys(styles)) {
				const expected = {};
				for (const step of naiveChain(styles, id).reverse())
					Object.assign(expected, styles[step]!.value);
				expect(resolve(id), `seed ${seed} ${id}`).toEqual(expected);
			}
			expect(inheritanceCycleEntries(styles)).toEqual(naiveCycleEntries(styles));
		}
	});

	it('loads a long basedOn chain in linear time', async () => {
		const load = async (count: number) => {
			const styles = Array.from(
				{ length: count },
				(_, index) =>
					`<w:style w:type="paragraph" w:styleId="s${index}">${index ? `<w:basedOn w:val="s${index - 1}"/>` : ''}<w:rPr><w:b/></w:rPr></w:style>`,
			).join('');
			const zip = await JSZip.loadAsync(await docx(p('text')));
			zip.file(
				'word/styles.xml',
				`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`,
			);
			const bytes = await zip.generateAsync({ type: 'uint8array' });
			return () => loadDocx(bytes);
		};
		const small = await load(3000);
		const large = await load(12000);
		// Four times the styles: about four times the time when linear, sixteen when quadratic.
		expect(await growthAsync(small, large)).toBeLessThan(8);
	}, 60_000);
});
