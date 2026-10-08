import { NS, buildXml, children, elements, first, parseXml, type XmlElement } from '../xml/index';
import { drawingColorXml } from './write-color';
import { parseDrawingColorIn } from './drawing-color';
import type { DrawingFill } from './types';

/** Replace only the DrawingML fill, retaining geometry, outlines, effects and extensions. */
export function setDrawingFillXml(properties: XmlElement, xml: string): void {
	const doc = properties.ownerDocument;
	const wrapper = parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`);
	const fill = elements(wrapper.documentElement)[0];
	const names = ['solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill', 'noFill'];
	if (
		elements(wrapper.documentElement).length !== 1 ||
		!fill ||
		fill.namespaceURI !== NS.a ||
		!names.includes(fill.localName)
	)
		throw new Error('Invalid DrawingML fill XML');
	const existing = elements(properties).filter(
		(node) => node.namespaceURI === NS.a && names.includes(node.localName),
	);
	const next = doc.importNode(fill, true);
	if (existing[0]) properties.replaceChild(next, existing[0]);
	else
		properties.insertBefore(
			next,
			elements(properties).find(
				(node) =>
					node.namespaceURI === NS.a &&
					['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].includes(node.localName),
			) ?? null,
		);
	for (const extra of existing.slice(1)) properties.removeChild(extra);
}

/** Serializes modelled DrawingML fills, retaining imported gradient flags and extensions. */
export function drawingFillXml(fill: DrawingFill): string | undefined {
	if (fill.kind === 'none') return '<a:noFill/>';
	if (fill.kind === 'solid') return `<a:solidFill>${drawingColorXml(fill.color)}</a:solidFill>`;
	if (fill.kind !== 'gradient') return undefined;
	const doc = parseXml(fill.sourceXml ?? `<a:gradFill xmlns:a="${NS.a}"/>`);
	const root = doc.documentElement;
	if (root.namespaceURI !== NS.a || root.localName !== 'gradFill')
		throw new Error('Invalid gradient source XML');
	const fragment = (xml: string) =>
		doc.importNode(elements(parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`).documentElement)[0]!, true);
	let list = first(root, 'gsLst', NS.a);
	if (!list) {
		list = doc.createElementNS(NS.a, 'a:gsLst');
		root.insertBefore(list, root.firstChild);
	}
	// Unknown imported stops are opaque preservation data, absent from the editable model.
	const stops = children(list, 'gs', NS.a).filter(
		(node) =>
			parseDrawingColorIn(node) &&
			Number.isFinite(Number.parseInt(node.getAttribute('pos') ?? '', 10)),
	);
	fill.stops.forEach((stop, index) => {
		const node = stops[index] ?? doc.createElementNS(NS.a, 'a:gs');
		if (!node.parentNode) list!.appendChild(node);
		node.setAttribute('pos', String(Math.round(stop.position * 1000)));
		for (const old of elements(node))
			if (old.namespaceURI === NS.a && old.localName.endsWith('Clr')) node.removeChild(old);
		node.insertBefore(fragment(drawingColorXml(stop.color)), node.firstChild);
	});
	for (const stop of stops.slice(fill.stops.length)) list.removeChild(stop);
	for (const node of elements(root))
		if (
			node.namespaceURI === NS.a &&
			((node.localName === 'lin' && (fill.angle === undefined || fill.path !== undefined)) ||
				(node.localName === 'path' && fill.path === undefined))
		)
			root.removeChild(node);
	if (fill.angle !== undefined && fill.path === undefined) {
		let linear = first(root, 'lin', NS.a);
		if (!linear) {
			linear = doc.createElementNS(NS.a, 'a:lin');
			root.insertBefore(
				linear,
				first(root, 'tileRect', NS.a) ?? first(root, 'extLst', NS.a) ?? null,
			);
		}
		linear.setAttribute('ang', String(Math.round(fill.angle * 60000)));
		if (fill.scaled !== undefined) linear.setAttribute('scaled', fill.scaled ? '1' : '0');
	}
	if (fill.path !== undefined) {
		let path = first(root, 'path', NS.a);
		if (!path) {
			path = doc.createElementNS(NS.a, 'a:path');
			root.insertBefore(path, first(root, 'tileRect', NS.a) ?? first(root, 'extLst', NS.a) ?? null);
		}
		path.setAttribute('path', fill.path);
		if (!fill.fillToRect) {
			const old = first(path, 'fillToRect', NS.a);
			if (old) path.removeChild(old);
		} else {
			let focus = first(path, 'fillToRect', NS.a);
			if (!focus) {
				focus = doc.createElementNS(NS.a, 'a:fillToRect');
				path.appendChild(focus);
			}
			for (const [key, value] of Object.entries(fill.fillToRect))
				focus.setAttribute(key, String(Math.round(value * 100000)));
		}
	}
	if (fill.tileRect) {
		let tile = first(root, 'tileRect', NS.a);
		if (!tile) {
			tile = doc.createElementNS(NS.a, 'a:tileRect');
			root.insertBefore(tile, first(root, 'extLst', NS.a) ?? null);
		}
		for (const [key, value] of Object.entries(fill.tileRect))
			tile.setAttribute(key, String(Math.round(value * 100000)));
	}
	return buildXml(doc);
}
