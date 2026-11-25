// Structural OPC + [MS-VSDX] package validator for .vsdx files.
// Validates ZIP/OPC structure, [Content_Types].xml, relationship integrity, [MS-VSDX] root
// elements and namespaces, strict XML well-formedness, page/shape ID uniqueness, and clean
// round-trip parsing.
// Usage: bun scripts/validate-vsdx.ts <file-or-dir> [...]
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, join, posix, resolve } from 'node:path';
import { parseContentTypes } from '../src/core/opc/content-types';
import { parseRelationships } from '../src/core/opc/relationships';
import { VisioPackage } from '../src/core/visio/package';
import { DEFAULTS } from '../src/core/visio/package-common';
import { parseVsdx } from '../src/core/visio/parser';
import { VISIO_NS } from '../src/core/visio/sheet';
import { inspectXml } from '../src/core/visio/xml-validation';
import { elements, NS, parseXml } from '../src/core/xml/index';

const VISIO_REL_DOC = 'http://schemas.microsoft.com/visio/2010/relationships/document';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

const EXPECTED_ROOTS: ReadonlyArray<readonly [RegExp, string]> = [
	[/^visio\/document\.xml$/i, 'VisioDocument'],
	[/^visio\/pages\/pages\.xml$/i, 'Pages'],
	[/^visio\/pages\/page\d+\.xml$/i, 'PageContents'],
	[/^visio\/windows\.xml$/i, 'Windows'],
	[/^visio\/masters\/masters\.xml$/i, 'Masters'],
	[/^visio\/masters\/master\d+\.xml$/i, 'MasterContents'],
];

function resolveRelTarget(relsPath: string, target: string): string {
	if (target.startsWith('/')) return posix.normalize(target.slice(1));
	const baseDir =
		relsPath === '_rels/.rels'
			? ''
			: posix.dirname(relsPath.replace(/\/_rels\/([^/]+)\.rels$/, '/$1'));
	return posix.normalize(baseDir ? posix.join(baseDir, target) : target);
}

/** Validates a single .vsdx archive and returns a list of errors (empty when valid). */
export async function validateVsdxBytes(bytes: Uint8Array): Promise<string[]> {
	const errors: string[] = [];
	let pkg: VisioPackage;
	try {
		pkg = await VisioPackage.open(bytes);
	} catch (error) {
		return [
			`Failed to open VSDX package: ${error instanceof Error ? error.message : String(error)}`,
		];
	}

	const allPaths = new Set(pkg.paths());
	const hasPath = (path: string) => allPaths.has(path);
	const decoder = new TextDecoder();
	const readText = async (path: string) => decoder.decode(await pkg.readBytes(path));

	// 1. [Content_Types].xml
	if (!hasPath('[Content_Types].xml')) {
		errors.push('Missing required part [Content_Types].xml');
	} else {
		try {
			const ctXml = await readText('[Content_Types].xml');
			inspectXml(ctXml, DEFAULTS, () => {});
			const doc = parseXml(ctXml);
			if (
				doc.documentElement.localName !== 'Types' ||
				doc.documentElement.namespaceURI !== CONTENT_TYPES_NS
			) {
				errors.push(
					`[Content_Types].xml has unexpected root <${doc.documentElement.localName}> in namespace "${doc.documentElement.namespaceURI ?? ''}"`,
				);
			}
			const ct = parseContentTypes(ctXml);
			if (!ct.defaults.has('rels'))
				errors.push('[Content_Types].xml is missing Default for extension "rels"');
			if (!ct.defaults.has('xml'))
				errors.push('[Content_Types].xml is missing Default for extension "xml"');
			const docType = ct.overrides.get('/visio/document.xml');
			if (
				docType !== 'application/vnd.ms-visio.drawing.main+xml' &&
				docType !== 'application/vnd.ms-visio.stencil.main+xml'
			) {
				errors.push(
					`[Content_Types].xml has unexpected Override for /visio/document.xml: "${docType ?? ''}"`,
				);
			}
			for (const path of allPaths) {
				if (/^visio\/pages\/page\d+\.xml$/i.test(path)) {
					const pageType = ct.overrides.get(`/${path}`);
					if (pageType !== 'application/vnd.ms-visio.page+xml') {
						errors.push(
							`[Content_Types].xml is missing page Override for /${path} (got "${pageType ?? ''}")`,
						);
					}
				}
			}
		} catch (error) {
			errors.push(
				`Invalid [Content_Types].xml: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
	}

	// 2. Package root relationships (_rels/.rels) and all internal .rels targets
	if (!hasPath('_rels/.rels')) {
		errors.push('Missing required package relationship part _rels/.rels');
	}
	for (const path of allPaths) {
		if (!path.endsWith('.rels')) continue;
		try {
			const xml = await readText(path);
			inspectXml(xml, DEFAULTS, () => {});
			const rels = parseRelationships(xml);
			if (path === '_rels/.rels') {
				const docRel = [...rels.values()].find((rel) => rel.type === VISIO_REL_DOC);
				if (!docRel) errors.push(`_rels/.rels is missing relationship ${VISIO_REL_DOC}`);
			}
			for (const [id, rel] of rels) {
				if (rel.mode === 'External') continue;
				const target = resolveRelTarget(path, rel.target);
				if (!hasPath(target)) {
					errors.push(`${path}: relationship ${id} points to missing part "${target}"`);
				}
			}
		} catch (error) {
			errors.push(`Invalid ${path}: ${error instanceof Error ? error.message : String(error)}`);
		}
	}

	// 3. Visio XML parts: strict XML, root element, [MS-VSDX] namespace, and part-level invariants
	if (!hasPath('visio/document.xml')) errors.push('Missing required part visio/document.xml');

	for (const path of allPaths) {
		const expected = EXPECTED_ROOTS.find(([pattern]) => pattern.test(path));
		if (!expected) continue;
		try {
			const xml = await readText(path);
			inspectXml(xml, DEFAULTS, () => {});
			const doc = parseXml(xml);
			const root = doc.documentElement;
			if (root.localName !== expected[1] || root.namespaceURI !== VISIO_NS) {
				errors.push(
					`${path}: expected <${expected[1]} xmlns="${VISIO_NS}">, got <${root.localName} xmlns="${root.namespaceURI ?? ''}">`,
				);
				continue;
			}

			if (/^visio\/pages\/pages\.xml$/i.test(path)) {
				const pageRels = hasPath('visio/pages/_rels/pages.xml.rels')
					? parseRelationships(await readText('visio/pages/_rels/pages.xml.rels'))
					: new Map();
				const seenIds = new Set<string>();
				for (const page of elements(root).filter((el) => el.localName === 'Page')) {
					const id = page.getAttribute('ID');
					if (!id) {
						errors.push(`${path}: <Page> is missing ID attribute`);
					} else if (seenIds.has(id)) {
						errors.push(`${path}: duplicate Page ID "${id}"`);
					} else {
						seenIds.add(id);
					}
					const relEl = elements(page).find((el) => el.localName === 'Rel');
					const relId = relEl?.getAttributeNS(NS.r, 'id') ?? relEl?.getAttribute('r:id');
					if (!relId || !pageRels.has(relId)) {
						errors.push(
							`${path}: Page ID="${id ?? '?'}" references missing relationship "${relId ?? ''}"`,
						);
					}
				}
			}

			if (/^visio\/pages\/page\d+\.xml$/i.test(path)) {
				const seenShapes = new Set<string>();
				const shapesContainer = elements(root).find((el) => el.localName === 'Shapes');
				const walkShapes = (parent: Element) => {
					for (const child of elements(parent)) {
						if (child.localName === 'Shape') {
							const id = child.getAttribute('ID');
							if (!id || !/^\d+$/.test(id)) {
								errors.push(`${path}: <Shape> has invalid or missing numeric ID "${id ?? ''}"`);
							} else if (seenShapes.has(id)) {
								errors.push(`${path}: duplicate Shape ID "${id}"`);
							} else {
								seenShapes.add(id);
							}
							const nested = elements(child).find((el) => el.localName === 'Shapes');
							if (nested) walkShapes(nested);
						}
					}
				};
				if (shapesContainer) walkShapes(shapesContainer);
			}
		} catch (error) {
			errors.push(`Invalid ${path}: ${error instanceof Error ? error.message : String(error)}`);
		}
	}

	// 4. Full parser round-trip
	try {
		const drawing = await parseVsdx(bytes);
		if (drawing.format !== 'vsdx') {
			errors.push(`parseVsdx reported unexpected format "${drawing.format}"`);
		}
	} catch (error) {
		errors.push(`parseVsdx failed: ${error instanceof Error ? error.message : String(error)}`);
	}

	return errors;
}

function collectVsdxFiles(inputs: readonly string[]): string[] {
	const files: string[] = [];
	const walk = (target: string) => {
		const stat = statSync(target);
		if (stat.isDirectory()) {
			for (const entry of readdirSync(target).sort()) walk(join(target, entry));
		} else if (target.toLowerCase().endsWith('.vsdx')) {
			files.push(target);
		}
	};
	for (const input of inputs) walk(resolve(input));
	return files;
}

async function main() {
	const inputs = process.argv.slice(2);
	if (!inputs.length) {
		console.error('Usage: bun scripts/validate-vsdx.ts <file-or-dir> [...]');
		process.exitCode = 2;
		return;
	}
	const files = collectVsdxFiles(inputs);
	if (!files.length) {
		console.error('No .vsdx files found to validate.');
		process.exitCode = 2;
		return;
	}
	let failed = 0;
	for (const file of files) {
		const bytes = new Uint8Array(readFileSync(file));
		const errors = await validateVsdxBytes(bytes);
		if (!errors.length) {
			console.log(`OK   ${basename(file)}`);
			continue;
		}
		failed++;
		console.error(`FAIL ${basename(file)} (${errors.length} error(s)):`);
		for (const error of errors) console.error(`  ${error}`);
	}
	console.log(
		`Validated ${files.length} VSDX package(s): ${files.length - failed} passed, ${failed} failed.`,
	);
	if (failed > 0) process.exitCode = 1;
}

if (import.meta.main) {
	main().catch((error) => {
		console.error(error);
		process.exitCode = 1;
	});
}
