import { renderSvgPreview } from '../../../drawingml/svg-preview';
import { renderWordArtSvg } from '../../../drawingml/word-art-svg';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'WordArt rejects a dimension supplied as text by a JavaScript caller',
		limits: { ms: 1_000, heapMb: 64 },
		run() {
			return Reflect.apply(renderWordArtSvg, undefined, [
				{
					id: 'preview',
					style: { fill: '26374A' },
					lines: ['Sample'],
					width: '160',
					height: 80,
					fontPx: 32,
				},
			]);
		},
		expected: '',
	},
	{
		name: 'gallery SVG previews reject links and styles outside the static preview subset',
		limits: { ms: 1_000, heapMb: 64 },
		run() {
			return {
				link: renderSvgPreview(32, 20, '', '<a><text>Details</text></a>'),
				style: renderSvgPreview(32, 20, '', '<rect style="filter:var(--preview)"/>'),
			};
		},
		expected: { link: '', style: '' },
	},
];
