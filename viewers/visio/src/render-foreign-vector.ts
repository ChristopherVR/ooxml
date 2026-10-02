import {
	validateVisioForeignVector,
	type VisioForeignVector,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorNode,
	type VisioForeignVectorCommand,
} from 'ooxml-core/visio';

const NS = 'http://www.w3.org/2000/svg';
let nextResource = 0;
const pathData = (commands: readonly VisioForeignVectorCommand[]) =>
	commands
		.map(({ command, values }) => `${command}${values.length ? ' ' + values.join(' ') : ''}`)
		.join(' ');
/** Render an independently validated inert scene. Never forwards source tags, IDs or attributes. */
export function renderForeignVector(source: unknown, owner: Document = document): SVGSVGElement {
	const scene: VisioForeignVector = validateVisioForeignVector(source);
	const element = <K extends keyof SVGElementTagNameMap>(tag: K) => owner.createElementNS(NS, tag);
	const root = element('svg'),
		defs = element('defs'),
		prefix = `visio-vector-${++nextResource}-`;
	root.setAttribute('viewBox', `0 0 ${scene.width} ${scene.height}`);
	root.setAttribute('width', String(scene.width));
	root.setAttribute('height', String(scene.height));
	root.setAttribute('preserveAspectRatio', 'none');
	root.setAttribute('overflow', 'hidden');
	const transform = (node: SVGElement, matrix: readonly number[]) =>
		node.setAttribute('transform', `matrix(${matrix.join(' ')})`);
	const clip = (node: SVGElement, index: number | undefined) => {
		if (index !== undefined) node.setAttribute('clip-path', `url(#${prefix}${index})`);
	};
	const clipPath = (item: VisioForeignVectorClipPath): SVGPathElement => {
		const path = element('path');
		path.setAttribute('d', pathData(item.commands));
		path.setAttribute('clip-rule', item.clipRule);
		transform(path, item.matrix);
		clip(path, item.clipIndex);
		return path;
	};
	scene.clips.forEach((item, index) => {
		const path = element('clipPath');
		path.id = `${prefix}${index}`;
		path.setAttribute('clipPathUnits', 'userSpaceOnUse');
		transform(path, item.matrix);
		clip(path, item.clipIndex);
		for (const child of item.items) path.append(clipPath(child));
		defs.append(path);
	});
	root.append(defs);
	const draw = (item: VisioForeignVectorNode): SVGElement => {
		if (item.kind === 'group') {
			const group = element('g');
			transform(group, item.matrix);
			clip(group, item.clipIndex);
			for (const child of item.items) group.append(draw(child));
			return group;
		}
		const path = element('path'),
			paint = item.paint;
		path.setAttribute('d', pathData(item.commands));
		transform(path, item.matrix);
		clip(path, item.clipIndex);
		for (const [name, value] of [
			['fill', paint.fill],
			['stroke', paint.stroke],
			['fill-rule', paint.fillRule],
			['stroke-width', paint.strokeWidth],
			['stroke-miterlimit', paint.strokeMiterlimit],
			['stroke-linecap', paint.strokeLinecap],
			['stroke-linejoin', paint.strokeLinejoin],
			['opacity', paint.opacity],
			['fill-opacity', paint.fillOpacity],
			['stroke-opacity', paint.strokeOpacity],
		] as const)
			path.setAttribute(name, String(value));
		return path;
	};
	for (const item of scene.items) root.append(draw(item));
	return root;
}
