// DrawingML outline writer. Patch changed fields while preserving authored tokens and extensions.
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import { drawingDashLength, parseDrawingLine } from './drawing-fill';
import { drawingFillXml } from './write-fill';
import { patchDrawingColorIn } from './patch-color';
import type { DrawingLine } from './types';

const FILLS = ['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'];
const ORDER = [
	...FILLS,
	'prstDash',
	'custDash',
	'round',
	'bevel',
	'miter',
	'headEnd',
	'tailEnd',
	'extLst',
];
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function importXml(host: XmlElement, xml: string): XmlElement {
	const wrapper = parseXml(`<w xmlns:a="${NS.a}">${xml}</w>`);
	return host.ownerDocument.importNode(elements(wrapper.documentElement)[0]!, true);
}

function dropDrawingDeclarations(node: XmlElement): void {
	if (node.getAttribute('xmlns:a') === NS.a) node.removeAttribute('xmlns:a');
	for (const child of elements(node)) dropDrawingDeclarations(child);
}

function insert(host: XmlElement, child: XmlElement): void {
	const rank = ORDER.indexOf(child.localName);
	host.insertBefore(
		child,
		elements(host).find(
			(node) => node.namespaceURI === NS.a && ORDER.indexOf(node.localName) > rank,
		) ?? null,
	);
}

function remove(host: XmlElement, names: readonly string[]): void {
	for (const child of elements(host))
		if (child.namespaceURI === NS.a && names.includes(child.localName)) host.removeChild(child);
}

/** Applies an outline, retaining advanced fields omitted by existing partial callers. */
export function patchDrawingLine(ln: XmlElement, line: DrawingLine): void {
	const before = line.sourceXml ? parseDrawingLine(ln) : undefined;
	const changed = (key: keyof DrawingLine) => !before || !same(before[key], line[key]);
	for (const [key, attribute] of [
		['widthEmu', 'w'],
		['cap', 'cap'],
		['compound', 'cmpd'],
		['alignment', 'algn'],
	] as const) {
		if (
			!changed(key) ||
			((key === 'compound' || key === 'alignment') && !line.sourceXml && !(key in line))
		)
			continue;
		const value = line[key];
		if (value === undefined) ln.removeAttribute(attribute);
		else ln.setAttribute(attribute, String(typeof value === 'number' ? Math.round(value) : value));
	}
	if (changed('fill')) {
		if (line.fill?.kind === 'picture')
			throw new Error('Picture fills are not supported on DrawingML outlines');
		const solid = first(ln, 'solidFill', NS.a);
		if (solid && line.fill?.kind === 'solid') patchDrawingColorIn(solid, line.fill.color);
		else {
			const xml = line.fill && drawingFillXml(line.fill);
			if (xml || !line.fill) remove(ln, FILLS);
			if (xml) {
				const child = importXml(ln, xml);
				dropDrawingDeclarations(child);
				insert(ln, child);
			}
		}
	}
	if (changed('dash') || changed('customDash')) patchDash(ln, line);
	if ((line.sourceXml || 'join' in line) && changed('join')) {
		const original = line.join && first(ln, line.join.kind, NS.a);
		const node = original?.cloneNode(true) as XmlElement | undefined;
		remove(ln, ['round', 'bevel', 'miter']);
		if (line.join) {
			const next = node ?? ln.ownerDocument.createElementNS(NS.a, `a:${line.join.kind}`);
			if (line.join.limit === undefined) next.removeAttribute('lim');
			else next.setAttribute('lim', String(line.join.limit));
			insert(ln, next);
		}
	}
	for (const key of ['headEnd', 'tailEnd'] as const) {
		if ((!line.sourceXml && !(key in line)) || !changed(key)) continue;
		const end = line[key];
		const old = first(ln, key, NS.a);
		if (!end) {
			if (old) ln.removeChild(old);
			continue;
		}
		const node = old ?? ln.ownerDocument.createElementNS(NS.a, `a:${key}`);
		for (const [property, attribute] of [
			['type', 'type'],
			['width', 'w'],
			['length', 'len'],
		] as const) {
			const value = end[property];
			if (value === undefined) node.removeAttribute(attribute);
			else node.setAttribute(attribute, value);
		}
		if (!old) insert(ln, node);
	}
}

function patchDash(ln: XmlElement, line: DrawingLine): void {
	if (line.dash === undefined && line.customDash === undefined) {
		// Existing partial callers clear preset dashes while retaining opaque custom dashes.
		remove(ln, line.sourceXml ? ['prstDash', 'custDash'] : ['prstDash']);
		return;
	}
	if (line.dash !== undefined) {
		remove(ln, ['prstDash', 'custDash']);
		const node = ln.ownerDocument.createElementNS(NS.a, 'a:prstDash');
		node.setAttribute('val', line.dash);
		insert(ln, node);
		return;
	}
	const existing = first(ln, 'custDash', NS.a);
	const custom = existing?.cloneNode(true) as XmlElement | undefined;
	const node = custom ?? ln.ownerDocument.createElementNS(NS.a, 'a:custDash');
	const stops = elements(node).filter(
		(stop) =>
			stop.namespaceURI === NS.a &&
			stop.localName === 'ds' &&
			drawingDashLength(stop.getAttribute('d')) !== undefined &&
			drawingDashLength(stop.getAttribute('sp')) !== undefined,
	);
	line.customDash?.forEach((stop, index) => {
		if (
			drawingDashLength(String(stop.dash)) === undefined ||
			drawingDashLength(String(stop.space)) === undefined
		)
			throw new Error('Invalid DrawingML custom dash length');
		const next =
			stops[index] ??
			(stop.sourceXml
				? importXml(ln, stop.sourceXml)
				: ln.ownerDocument.createElementNS(NS.a, 'a:ds'));
		next.setAttribute('d', String(stop.dash));
		next.setAttribute('sp', String(stop.space));
		if (!next.parentNode) node.appendChild(next);
	});
	for (const stop of stops.slice(line.customDash?.length ?? 0)) node.removeChild(stop);
	remove(ln, ['prstDash', 'custDash']);
	insert(ln, node);
}

/** A new or source-preserving `a:ln`. */
export function drawingLineXml(line: DrawingLine): string {
	const doc = parseXml(line.sourceXml ?? `<a:ln xmlns:a="${NS.a}"/>`);
	if (doc.documentElement.namespaceURI !== NS.a || doc.documentElement.localName !== 'ln')
		throw new Error('Invalid DrawingML outline source XML');
	patchDrawingLine(doc.documentElement, line);
	return buildXml(doc);
}
