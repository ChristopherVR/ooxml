#!/usr/bin/env node
/**
 * Regenerate `diagram/builtin/style-{catalog,data}.ts` and `diagram/catalogue-native.ts`: every
 * built-in SmartArt colour (`dgm:colorsDef`) and quick-style (`dgm:styleDef`) definition, taken
 * from the parts PowerPoint wrote into `diagram/__fixtures__/smartart-definitions` (produced by
 * `make-smartart-definitions.ps1` through COM; see PROVENANCE.md).
 *
 * `style-data.ts` holds the definitions as one base64 blob of per-definition gzip members, so
 * applying one inflates only that definition; `style-catalog.ts` is the index (id and byte range).
 * `catalogue-native.ts` is the small gallery metadata the main `diagram` entry carries (name,
 * category, gallery order and what a gallery tile needs to draw the entry), so a gallery renders
 * without loading the definitions. Output is deterministic for the same fixture.
 *
 * Usage: `node scripts/diagram/gen-smartart-style-definitions.mjs` from `src/core`.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

import JSZip from 'jszip';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURE = path.resolve(HERE, '../../diagram/__fixtures__/smartart-definitions');
const BUILTIN = path.resolve(HERE, '../../diagram/builtin');
const DIAGRAM = path.resolve(HERE, '../../diagram');

const manifest = JSON.parse(
	readFileSync(path.join(FIXTURE, 'manifest.json'), 'utf8').replace(/^﻿/u, ''),
);
const zip = await JSZip.loadAsync(readFileSync(path.join(FIXTURE, 'smartart-definitions.pptx')));

async function part(name) {
	const file = zip.file(name);
	if (!file) throw new Error(`Missing part ${name}`);
	return file.async('string');
}

/** The diagram part of `kind` the graphic frame on slide `n` points at. */
async function definitionOfSlide(n, kind) {
	const rels = await part(`ppt/slides/_rels/slide${n}.xml.rels`);
	const type = kind === 'colors' ? 'diagramColors' : 'diagramQuickStyle';
	const target = new RegExp(
		`Target="([^"]+)"[^>]*Type="[^"]*/${type}"|Type="[^"]*/${type}"[^>]*Target="([^"]+)"`,
		'u',
	).exec(rels);
	const ref = target?.[1] ?? target?.[2];
	if (!ref) throw new Error(`Slide ${n} has no ${type} relationship`);
	return part(path.posix.normalize(path.posix.join('ppt/slides', ref)));
}

const strip = (xml) =>
	xml
		.replace(/^<\?xml[^>]*\?>\s*/u, '')
		.replace(/>\s+</gu, '><')
		.trim();

function schemeClrs(xml) {
	return [...xml.matchAll(/<a:schemeClr val="(\w+)"/gu)].map((match) => match[1]);
}

function labelXml(xml, name) {
	return new RegExp(`<dgm:styleLbl name="${name}">(.*?)</dgm:styleLbl>`, 'u').exec(xml)?.[1] ?? '';
}

/** Tone of a colour list entry: a shade as a negative, a tint as a positive fraction. */
function tones(list) {
	return [...list.matchAll(/<a:schemeClr val="\w+"(?:\/>|>(.*?)<\/a:schemeClr>)/gu)].map(
		(match) => {
			const shade = /<a:shade val="(\d+)"/u.exec(match[1] ?? '')?.[1];
			const tint = /<a:tint val="(\d+)"/u.exec(match[1] ?? '')?.[1];
			if (shade) return -Math.round((1 - Number(shade) / 100000) * 100) / 100;
			if (tint) return Math.round((1 - Number(tint) / 100000) * 100) / 100;
			return 0;
		},
	);
}

/** What a gallery tile needs to suggest a colour definition: node fills and lines. */
function colorPreview(xml) {
	const node = labelXml(xml, 'node1');
	const fill = /<dgm:fillClrLst([^>]*)>(.*?)<\/dgm:fillClrLst>/u.exec(node);
	const line = /<dgm:linClrLst[^>]*>(.*?)<\/dgm:linClrLst>/u.exec(node)?.[1] ?? '';
	// `span` is the schema default of `meth`.
	const meth = /meth="(\w+)"/u.exec(fill?.[1] ?? '')?.[1] ?? 'span';
	const fills = schemeClrs(fill?.[2] ?? '');
	const fillTones = tones(fill?.[2] ?? '');
	const transparent = /<a:alpha val="(\d+)"/u.test(fill?.[2] ?? '');
	return {
		fills,
		...(fillTones.some(Boolean) ? { tones: fillTones } : {}),
		line: schemeClrs(line)[0] ?? 'lt1',
		mode: meth,
		...(transparent ? { transparent: true } : {}),
	};
}

/** What a gallery tile needs to suggest a quick style: effect strength, outline, 3D. */
function stylePreview(xml) {
	const node = labelXml(xml, 'node1');
	const idx = (ref) => Number(new RegExp(`<a:${ref} idx="(\\d+)"`, 'u').exec(node)?.[1] ?? '0');
	const camera = /<a:camera prst="(\w+)"/u.exec(xml)?.[1] ?? 'orthographicFront';
	const bevel = /<a:bevelT\b/u.test(node);
	return {
		line: idx('lnRef'),
		fill: idx('fillRef'),
		effect: idx('effectRef'),
		...(camera !== 'orthographicFront' ? { camera } : {}),
		...(bevel ? { bevel: true } : {}),
	};
}

const definitions = [];
const seen = new Set();
for (const entry of manifest.entries) {
	const xml = strip(await definitionOfSlide(entry.slide, entry.kind));
	const id = /\buniqueId="([^"]+)"/u.exec(xml)?.[1];
	if (id !== entry.id) throw new Error(`Slide ${entry.slide}: ${id} is not ${entry.id}`);
	if (seen.has(id)) continue;
	seen.add(id);
	definitions.push({ ...entry, xml });
}

const members = [];
const index = [];
let offset = 0;
for (const definition of definitions) {
	const packed = gzipSync(Buffer.from(definition.xml, 'utf8'), { level: 9, mtime: 0 });
	members.push(packed);
	index.push({ id: definition.id, kind: definition.kind, offset, length: packed.length });
	offset += packed.length;
}

const header = '// Generated by scripts/diagram/gen-smartart-style-definitions.mjs. Do not edit.\n';
writeFileSync(
	path.join(BUILTIN, 'style-catalog.ts'),
	`${header}import type { BuiltinSmartArtDefinitionEntry } from './types';\n\n` +
		`export const BUILTIN_SMARTART_STYLE_DEFINITIONS: readonly BuiltinSmartArtDefinitionEntry[] = ${JSON.stringify(index, null, '\t')};\n`,
);
writeFileSync(
	path.join(BUILTIN, 'style-data.ts'),
	`${header}/** Per-definition gzip members, concatenated and base64 encoded (ranges in style-catalog.ts). */\n` +
		`export const BUILTIN_SMARTART_STYLE_DATA = '${Buffer.concat(members).toString('base64')}';\n`,
);

const short = (id) => id.split('/').at(-1);
const colors = definitions
	.filter((definition) => definition.kind === 'colors')
	.map((definition) => ({
		id: short(definition.id),
		uniqueId: definition.id,
		name: definition.name,
		category: definition.category,
		preview: colorPreview(definition.xml),
	}));
const styles = definitions
	.filter((definition) => definition.kind === 'style')
	.map((definition) => ({
		id: short(definition.id),
		uniqueId: definition.id,
		name: definition.name,
		category: definition.category,
		preview: stylePreview(definition.xml),
	}));
writeFileSync(
	path.join(DIAGRAM, 'catalogue-native.ts'),
	`${header}// The gallery metadata of Office's built-in SmartArt colour and quick-style definitions, in\n` +
		`// gallery order (PowerPoint ${manifest.powerpoint} build ${manifest.build}). The definitions themselves are in\n` +
		`// \`ooxml-core/diagram/builtin\` (\`loadBuiltinSmartArtStyleDefinitionXml\`).\n` +
		`import type { DiagramNativeColorEntry, DiagramNativeStyleEntry } from './catalogue-native-types';\n\n` +
		`export const DIAGRAM_NATIVE_COLOR_ENTRIES: readonly DiagramNativeColorEntry[] = ${JSON.stringify(colors, null, '\t')};\n\n` +
		`export const DIAGRAM_NATIVE_STYLE_ENTRIES: readonly DiagramNativeStyleEntry[] = ${JSON.stringify(styles, null, '\t')};\n`,
);
console.log(
	`${colors.length} colour and ${styles.length} quick-style definitions, ${offset} gzip bytes`,
);
