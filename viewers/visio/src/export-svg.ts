import {
	getVisioPageLayers,
	type VisioDocument,
	type VisioPage,
	type VisioShape,
} from 'ooxml-core/visio';
import { compatibilityNotes, compatibilityText, type CompatibilityNote } from './diagnostics.js';
import { renderPage, svgElement } from './render-svg.js';
import { assertViewableDocument } from './scene-validation.js';

export const MAX_SVG_EXPORT_BYTES = 16 * 1024 * 1024;
export interface SvgExportOptions {
	/** UTF-8 output ceiling. May lower, but never raise, the 16 MiB safety limit. */
	maxBytes?: number;
}
export interface SvgExportResult {
	readonly svg: string;
	readonly byteLength: number;
	readonly pageIndex: number;
	readonly pageName: string;
	/** Saved page dimensions in inches, independent of viewport zoom. */
	readonly width: number;
	readonly height: number;
	readonly diagnostics: CompatibilityNote[];
}
const APPROXIMATION =
	'Static SVG export is a rendered approximation, not an editable VSDX round-trip or a guarantee of Microsoft Visio fidelity. Fonts are not embedded.';
const INVALID_XML =
	'Invalid XML characters were replaced with the Unicode replacement character in the SVG export.';
const limitError = () =>
	new Error('This page exceeds the safe SVG export byte or amplification limit.');

/**
 * Export one page through the shared renderer, including its resolved backgrounds.
 * Requires a browser DOM at call time, never uploads, downloads or mutates viewer state.
 * Conservative preflight limits can reject a page whose eventual serialization would be smaller.
 */
export function exportPageSvg(
	model: VisioDocument,
	pageIndex = 0,
	options: SvgExportOptions = {},
): SvgExportResult {
	const maxBytes = options.maxBytes ?? MAX_SVG_EXPORT_BYTES;
	if (!Number.isSafeInteger(maxBytes) || maxBytes < 1 || maxBytes > MAX_SVG_EXPORT_BYTES)
		throw new Error('SVG maxBytes must be a positive integer no greater than 16 MiB.');
	assertViewableDocument(model);
	if (!Number.isSafeInteger(pageIndex) || pageIndex < 0 || pageIndex >= model.pages.length)
		throw new Error('The SVG export page index is out of range.');
	const page = model.pages[pageIndex]!;
	preflight(model, page, maxBytes);
	if (typeof document === 'undefined' || typeof XMLSerializer === 'undefined')
		throw new Error('SVG export requires a browser DOM and XMLSerializer.');
	const rendered = renderPage(model, page, { static: true });
	try {
		const { svg } = rendered;
		svg.setAttribute('width', `${page.width}in`);
		svg.setAttribute('height', `${page.height}in`);
		// XML 1.0 cannot represent isolated surrogates or control characters. Keep all valid Unicode.
		let replaced = normalizeXmlTree(svg);
		const notes = compatibilityNotes(model.diagnostics, [APPROXIMATION, ...rendered.warnings]);
		const pageName = xmlText(page.name);
		for (const note of notes) {
			const message = xmlText(note.message);
			if (message !== note.message) replaced = true;
			note.message = message;
		}
		if (replaced) notes.push({ message: INVALID_XML, severity: 'warning', count: 1 });
		const description = svgElement('desc');
		description.textContent = notes.map(compatibilityText).join('\n');
		svg.insertBefore(description, svg.children[1] ?? null);
		const metadata = svgElement('metadata');
		metadata.textContent = JSON.stringify({
			format: 'visio-viewer-static-svg',
			page: {
				index: pageIndex,
				name: pageName,
				width: page.width,
				height: page.height,
				unit: 'in',
			},
			diagnostics: notes,
		});
		svg.append(metadata);
		// Count the DOM before allocating the serialized string, including escaped attribute values.
		assertSerializedBudget(svg, maxBytes);
		const output = new XMLSerializer().serializeToString(svg);
		const byteLength = new TextEncoder().encode(output).byteLength;
		if (byteLength > maxBytes) throw limitError();
		return {
			svg: output,
			byteLength,
			pageIndex,
			pageName,
			width: page.width,
			height: page.height,
			diagnostics: notes,
		};
	} finally {
		rendered.dispose();
	}
}

/** Bound markup, raster encoding and worst-case text-fragment expansion before allocating DOM. */
function preflight(model: VisioDocument, page: VisioPage, maxBytes: number): void {
	let remaining = maxBytes;
	const reserve = (bytes: number) => {
		remaining -= bytes;
		if (remaining < 0) throw limitError();
	};
	reserve(8192 + page.name.length * 36);
	for (const note of model.diagnostics) reserve(384 + note.message.length * 36);
	const images = new Map<Uint8Array, Set<string>>();
	const stack: VisioShape[] = [];
	for (const layer of getVisioPageLayers(model, page.id)) stack.push(...layer.shapes);
	while (stack.length) {
		const shape = stack.pop()!;
		if (shape.hidden) continue;
		stack.push(...shape.children);
		reserve(1024 + (shape.text.plainText || shape.name).length * 6);
		if (shape.kind === 'group' && shape.groupDisplayMode === 0) continue;
		for (const geometry of shape.geometry)
			reserve(
				4096 +
					geometry.path.length * 6 +
					shape.style.lineColor.length * 24 +
					shape.style.fill.length * 6,
			);
		for (const stop of shape.style.fillGradient?.stops ?? []) reserve(256 + stop.color.length * 6);
		if (shape.image) {
			reserve(1536);
			const image = shape.image;
			const types = images.get(image.bytes) ?? new Set<string>();
			if (!types.has(image.mimeType)) {
				reserve(512 + 4 * Math.ceil(image.bytes.byteLength / 3));
				types.add(image.mimeType);
				images.set(image.bytes, types);
			}
		}
		if (!shape.text.plainText) continue;
		const text = shape.text;
		const runs = text.runs.length
			? text.runs
			: [{ text: text.plainText, fontFamily: text.fontFamily, color: text.color }];
		// Every character could be wrapped separately. Repeating long font/style values must be bounded.
		let characters = 0;
		for (const run of runs) {
			characters += run.text.length;
			reserve(run.text.length * (512 + run.fontFamily.length * 6 + run.color.length * 6));
		}
		const lineCount = Math.min(5000, characters + (text.paragraphs?.length ?? 1));
		reserve(lineCount * (640 + (text.backgroundColor ? 512 + text.backgroundColor.length * 6 : 0)));
		for (const paragraph of text.paragraphs ?? [])
			if (paragraph.bullet)
				reserve(
					512 +
						(paragraph.bullet.text.length +
							paragraph.bullet.fontFamily.length +
							text.color.length) *
							6,
				);
	}
}

function xmlText(value: string): string {
	return value.replace(
		/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF\uD800-\uDFFF]/gu,
		'\uFFFD',
	);
}
function normalizeXmlTree(root: Element): boolean {
	let replaced = false;
	const stack: Node[] = [root];
	while (stack.length) {
		const node = stack.pop()!;
		if (node.nodeType === 1)
			for (const attribute of Array.from((node as Element).attributes)) {
				const normalized = xmlText(attribute.value);
				if (normalized !== attribute.value) {
					replaced = true;
					attribute.value = normalized;
				}
			}
		else if (node.nodeType === 3) {
			const normalized = xmlText(node.nodeValue ?? '');
			if (normalized !== node.nodeValue) {
				replaced = true;
				node.nodeValue = normalized;
			}
		}
		stack.push(...node.childNodes);
	}
	return replaced;
}

/** An upper bound, without allocating a second escaped copy of large paths or data URLs. */
function assertSerializedBudget(root: Element, maxBytes: number): void {
	let bytes = 0;
	const add = (value: string) => {
		for (const point of value) {
			const code = point.codePointAt(0)!;
			bytes += /[&<>"'\t\n\r]/u.test(point)
				? 6
				: code < 0x80
					? 1
					: code < 0x800
						? 2
						: code < 0x10000
							? 3
							: 4;
			if (bytes > maxBytes) throw limitError();
		}
	};
	const stack: Node[] = [root];
	while (stack.length) {
		const node = stack.pop()!;
		if (node.nodeType === 1) {
			const element = node as Element;
			bytes += element.tagName.length * 2 + 5;
			for (const attribute of element.attributes) {
				bytes += attribute.name.length + 4;
				add(attribute.value);
			}
		} else if (node.nodeType === 3) add(node.nodeValue ?? '');
		if (bytes > maxBytes) throw limitError();
		stack.push(...node.childNodes);
	}
}
