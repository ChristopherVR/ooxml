import { applyPictureLayout } from './picture-layout';
import { applyPictureCrop } from './picture-crop';
import { applyPictureEffects } from './picture-effects';
import { NS, buildXml, first, parseXml } from '../xml/index';
import type { DrawingFill } from './types';
import { patchDrawingColorIn } from './patch-color';
import { parseDrawingFill } from './drawing-fill';

/** Pattern colours stay authored theme tokens; picture relationships and bytes belong to the host. */
export function drawingPatternPictureXml(
	fill: Extract<DrawingFill, { kind: 'pattern' | 'picture' }>,
): string {
	const name = fill.kind === 'pattern' ? 'pattFill' : 'blipFill';
	const doc = parseXml(
		fill.sourceXml ??
			`<a:${name} xmlns:a="${NS.a}"${fill.kind === 'picture' ? ` xmlns:r="${NS.r}"` : ''}/>`,
	);
	const root = doc.documentElement;
	if (root.namespaceURI !== NS.a || root.localName !== name)
		throw new Error('Invalid DrawingML fill source XML');
	const holder = doc.createElementNS(NS.a, 'a:spPr');
	holder.appendChild(root.cloneNode(true));
	const before = fill.sourceXml ? parseDrawingFill(holder) : undefined;
	if (fill.kind === 'pattern') {
		if (!fill.preset && !fill.sourceXml) throw new Error('A DrawingML pattern needs a preset');
		if (before?.kind !== 'pattern' || before.preset !== fill.preset)
			root.setAttribute('prst', fill.preset);
		for (const [key, tag] of [
			['foreground', 'fgClr'],
			['background', 'bgClr'],
		] as const) {
			const color = fill[key];
			if (before?.kind === 'pattern' && JSON.stringify(before[key]) === JSON.stringify(color))
				continue;
			const old = first(root, tag, NS.a);
			if (!color) {
				if (old) root.removeChild(old);
				continue;
			}
			const node = old ?? doc.createElementNS(NS.a, `a:${tag}`);
			patchDrawingColorIn(node, color);
			if (!old)
				root.insertBefore(
					node,
					tag === 'fgClr'
						? (first(root, 'bgClr', NS.a) ?? root.firstChild)
						: (first(root, 'extLst', NS.a) ?? null),
				);
		}
	} else {
		if (!fill.sourceXml && !fill.relId && !fill.linkRelId)
			throw new Error('A DrawingML picture fill needs a host relationship ID');
		let blip = first(root, 'blip', NS.a);
		if (!blip && fill.sourceXml && !fill.relId && !fill.linkRelId) return buildXml(doc);
		if (!blip) {
			blip = doc.createElementNS(NS.a, 'a:blip');
			root.insertBefore(blip, root.firstChild);
		}
		for (const [key, attribute] of [
			['relId', 'embed'],
			['linkRelId', 'link'],
		] as const) {
			const value = fill[key];
			if (before?.kind === 'picture' && before[key] === value) continue;
			if (value !== undefined && (!value || /\s/u.test(value)))
				throw new Error('Invalid DrawingML picture relationship ID');
			if (value === undefined) blip.removeAttributeNS(NS.r, attribute);
			else if (blip.getAttributeNS(NS.r, attribute) !== value)
				blip.setAttributeNS(NS.r, `r:${attribute}`, value);
		}
		if (
			fill.crop !== undefined &&
			JSON.stringify(before?.kind === 'picture' ? before.crop : undefined) !==
				JSON.stringify(fill.crop)
		)
			applyPictureCrop(root, fill.crop);
		if (
			fill.effects !== undefined &&
			JSON.stringify(before?.kind === 'picture' ? (before.effects ?? []) : []) !==
				JSON.stringify(fill.effects ?? [])
		)
			applyPictureEffects(blip, fill.effects ?? []);
		if (fill.layout !== undefined) applyPictureLayout(root, fill.layout);
		if (!fill.sourceXml && !fill.layout?.tile && !first(root, 'stretch', NS.a)) {
			const stretch = doc.createElementNS(NS.a, 'a:stretch');
			stretch.appendChild(doc.createElementNS(NS.a, 'a:fillRect'));
			root.appendChild(stretch);
		}
	}
	return buildXml(doc);
}
