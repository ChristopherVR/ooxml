import type { ShapeNode } from 'ooxml-core/xlsx';
import type { EquationPreview } from '../../math/equation-render';

type Paragraphs = NonNullable<ShapeNode['equationParagraphs']>;
type Segment = Paragraphs[number][number];
const previews = new WeakMap<
	Segment,
	{ omml: string; pending: Promise<EquationPreview | undefined> }
>();

/** Mounts into the SVG shape's text rectangle so rotation, scaling and grouping follow the shape. */
export function mountShapeEquations(
	svg: Element,
	data: readonly { paragraphs: Paragraphs; scale: number }[],
	label: string,
): void {
	const doc = svg.ownerDocument;
	for (const placeholder of svg.querySelectorAll('foreignObject[data-equation-text]')) {
		const item = data[Number(placeholder.getAttribute('data-equation-text'))];
		if (!item) continue;
		const content = doc.createElement('div');
		content.style.fontSize = `${(44 / 3) * item.scale}px`;
		content.style.whiteSpace = 'pre-wrap';
		for (const paragraph of item.paragraphs) {
			const row = doc.createElement('div');
			for (const segment of paragraph) {
				if (segment.kind === 'text') {
					row.append(segment.text);
					continue;
				}
				const span = doc.createElement('span');
				span.textContent = segment.fallback || label;
				row.append(span);
				const omml = segment.omml;
				if (!omml) continue;
				const renderer = import('../../math/equation-render');
				let cached = previews.get(segment);
				if (cached?.omml !== omml) {
					cached = { omml, pending: renderer.then(({ equationPreview }) => equationPreview(omml)) };
					previews.set(segment, cached);
				}
				void Promise.all([cached.pending, renderer]).then(
					([preview, { renderEquation }]) => {
						const math = preview && renderEquation(doc, preview, false, label);
						if (math) span.replaceChildren(math);
					},
					() => undefined,
				);
			}
			content.append(row);
		}
		placeholder.append(content);
	}
}
