/**
 * Assembles a drawing: the core library creates the package and its pages and applies the
 * structural edits (containers, swimlanes, shape data, hyperlinks, layers, data links); this
 * module adds the parts the library has no writer for (styles, masters, drawn shapes, routes).
 */
import { createVsdx } from '../../../../src/core/visio/create-document';
import { editVsdx } from '../../../../src/core/visio/edit';
import type { VisioEdit } from '../../../../src/core/visio/edit-commands';
import { patchPackage, readPackage } from '../../lib/package';
import { SAMPLE_DATE_ISO } from '../../lib/types';
import { type MasterDef, Stencil } from './masters';
import { PageBuilder } from './page';
import {
	APP_PROPERTIES_XML,
	EMPTY_RELS,
	PROPERTY_CONTENT_TYPES,
	corePropertiesXml,
	insertBefore,
	withContentTypes,
	withPropertyRelationships,
	withRelationship,
} from './parts';
import { type RecordsetSpec, rowsetXml } from './recordset';
import { type StyleDef, StyleBook } from './styles';
import { textBody } from './text';

export interface PageSpec {
	name: string;
	background?: boolean;
	/** Name of the background page shown behind this page. */
	backPage?: string;
}

export interface DrawingSpec {
	width: number;
	height: number;
	/** Default text and line colour. */
	ink: string;
	styles: readonly StyleDef[];
	masters: readonly MasterDef[];
	pages: readonly PageSpec[];
	title: string;
	subject: string;
	description: string;
	keywords: readonly string[];
	/** The theme part and the PageSheet cells that select it, when the drawing has a theme. */
	theme?: { xml: string; pageCells: string };
}

const decoder = new TextDecoder();
/** Written as a code point so that this file holds none itself. */
const EM_DASH = String.fromCodePoint(0x2014);
const MASTER_REL = 'http://schemas.microsoft.com/visio/2010/relationships/master';
/** Fonts a sample must never name (the brief's rule 5). */
const FORBIDDEN_FONTS =
	/Calibri|Aptos|Segoe|Arial|Times New Roman|Cambria|Consolas|Georgia|Verdana|Tahoma|Courier New|Symbol|Wingdings/u;

export class Drawing {
	bytes: Uint8Array;
	readonly styles: StyleBook;
	readonly stencil: Stencil;
	/** Diagnostic codes the library reported while editing, for the sample's report. */
	readonly notes = new Set<string>();
	private readonly builders = new Map<string, PageBuilder>();
	private readonly pageParts = new Map<string, string>();
	private readonly layerNames = new Map<string, string[]>();
	private readonly recordsets: RecordsetSpec[] = [];

	private constructor(
		bytes: Uint8Array,
		private readonly spec: DrawingSpec,
	) {
		this.bytes = bytes;
		this.styles = new StyleBook(spec.styles, spec.ink, spec.theme?.pageCells);
		this.stencil = new Stencil(spec.masters);
	}

	/** The package with its pages, styles and masters, ready for shapes. */
	static async create(spec: DrawingSpec): Promise<Drawing> {
		const drawing = new Drawing(await createVsdx({ width: spec.width, height: spec.height }), spec);
		await drawing.createPages();
		await drawing.writeDocumentParts();
		return drawing;
	}

	private async createPages(): Promise<void> {
		const { pages } = this.spec;
		const first = pages[0];
		if (!first) throw new Error('A drawing needs a page.');
		const idOf = (name: string) => {
			const index = pages.findIndex((page) => page.name === name);
			if (index < 0) throw new Error(`Unknown page: ${name}`);
			return String(index);
		};
		await this.edit([{ type: 'rename-page', pageId: '0', name: first.name }], 'rename page');
		for (let index = 1; index < pages.length; index++) {
			const page = pages[index];
			if (!page) continue;
			await this.edit(
				[
					{
						type: 'insert-page',
						pageId: String(index),
						afterPageId: String(index - 1),
						name: page.name,
					},
				],
				`insert page ${page.name}`,
			);
		}
		for (const page of pages.filter((entry) => entry.background))
			await this.edit(
				[{ type: 'set-page-properties', pageId: idOf(page.name), background: true }],
				`background ${page.name}`,
			);
		for (const page of pages) {
			if (!page.backPage) continue;
			await this.edit(
				[{ type: 'set-page-properties', pageId: idOf(page.name), backPageId: idOf(page.backPage) }],
				`back page of ${page.name}`,
			);
		}
		await this.indexPages();
		pages.forEach((page, index) => {
			this.builders.set(
				page.name,
				new PageBuilder(String(index), page.name, this.styles, this.stencil),
			);
		});
	}

	/** Maps each page ID to its part through the pages part and its relationships. */
	private async indexPages(): Promise<void> {
		const pagesXml = await this.text('visio/pages/pages.xml');
		const rels = await this.text('visio/pages/_rels/pages.xml.rels');
		for (const match of pagesXml.matchAll(/<Page\b[^>]*\bID="(\d+)"[\s\S]*?<\/Page>/gu)) {
			const relId = /<Rel\b[^>]*r:id="([^"]+)"/u.exec(match[0])?.[1];
			const target = new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`, 'u').exec(rels)?.[1];
			const reversed = new RegExp(`Target="([^"]+)"[^>]*Id="${relId}"`, 'u').exec(rels)?.[1];
			const part = target ?? reversed;
			if (!match[1] || !part) throw new Error('A page has no part.');
			this.pageParts.set(match[1], `visio/pages/${part}`);
		}
	}

	private async writeDocumentParts(): Promise<void> {
		const { stencil, styles, spec } = this;
		const changes: Record<string, string> = {
			'visio/document.xml': styles.documentXml(),
			// A master resolves theme formulas against its own page sheet, so it selects the theme too.
			'visio/masters/masters.xml': stencil.mastersXml(spec.theme?.pageCells ?? ''),
			'visio/masters/_rels/masters.xml.rels': stencil.relsXml(),
		};
		const types: [string, string][] = [
			['/visio/masters/masters.xml', 'application/vnd.ms-visio.masters+xml'],
		];
		for (const master of stencil.masters) {
			changes[`visio/masters/master${master.id}.xml`] = stencil.masterXml(master, styles);
			types.push([`/visio/masters/master${master.id}.xml`, 'application/vnd.ms-visio.master+xml']);
		}
		let documentRels = withRelationship(
			await this.text('visio/_rels/document.xml.rels'),
			'rIdMasters',
			'http://schemas.microsoft.com/visio/2010/relationships/masters',
			'masters/masters.xml',
		);
		if (spec.theme) {
			changes['visio/theme/theme1.xml'] = spec.theme.xml;
			types.push([
				'/visio/theme/theme1.xml',
				'application/vnd.openxmlformats-officedocument.theme+xml',
			]);
			documentRels = withRelationship(
				documentRels,
				'rIdTheme',
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				'theme/theme1.xml',
			);
			const pagesXml = await this.text('visio/pages/pages.xml');
			changes['visio/pages/pages.xml'] = pagesXml.replaceAll(
				/(<PageSheet\b[^>]*>)/gu,
				`$1${spec.theme.pageCells}`,
			);
		}
		changes['visio/_rels/document.xml.rels'] = documentRels;
		changes['[Content_Types].xml'] = withContentTypes(
			await this.text('[Content_Types].xml'),
			types,
		);
		this.bytes = await patchPackage(this.bytes, changes);
	}

	async text(part: string): Promise<string> {
		const data = (await readPackage(this.bytes)).get(part);
		if (!data) throw new Error(`The package has no part ${part}.`);
		return decoder.decode(data);
	}

	async patch(changes: Record<string, string | Uint8Array | undefined>): Promise<void> {
		this.bytes = await patchPackage(this.bytes, changes);
	}

	page(name: string): PageBuilder {
		const builder = this.builders.get(name);
		if (!builder) throw new Error(`Unknown page: ${name}`);
		return builder;
	}

	pagePart(pageId: string): string {
		const part = this.pageParts.get(pageId);
		if (!part) throw new Error(`Unknown page ID: ${pageId}`);
		return part;
	}

	/** Runs one batch of library edits; a refusal names the step. */
	async edit(edits: readonly VisioEdit[], label: string): Promise<void> {
		if (edits.length === 0) return;
		try {
			const result = await editVsdx(this.bytes, [...edits]);
			this.bytes = result.bytes;
			for (const diagnostic of result.diagnostics) this.notes.add(diagnostic.code);
		} catch (error) {
			const code = (error as { code?: string }).code ?? '';
			throw new Error(`${label}: ${code} ${(error as Error).message}`, { cause: error });
		}
	}

	/**
	 * Writes the shapes built so far into the page part, then lets the library add their shape
	 * data, hyperlinks and layers. May be called again after more shapes were built.
	 */
	async place(page: PageBuilder): Promise<void> {
		const { shapes, connects } = page.flush();
		const part = this.pagePart(page.pageId);
		let xml = await this.text(part);
		xml = xml.replace(/<PageContents\b([^>]*)\/>/u, '<PageContents$1></PageContents>');
		xml = xml.replace(/<Shapes\s*\/>/u, '<Shapes></Shapes>');
		if (!xml.includes('</Shapes>'))
			xml = xml.replace(/(<PageContents\b[^>]*>)/u, '$1<Shapes></Shapes>');
		xml = insertBefore(xml, '</Shapes>', shapes);
		if (connects) {
			xml = xml.replace(/<Connects\s*\/>/u, '<Connects></Connects>');
			xml = xml.includes('</Connects>')
				? insertBefore(xml, '</Connects>', connects)
				: insertBefore(xml, '</PageContents>', `<Connects>${connects}</Connects>`);
		}
		const relsPart = part.replace(/([^/]+)$/u, '_rels/$1.rels');
		const existing = (await readPackage(this.bytes)).get(relsPart);
		let rels = existing ? decoder.decode(existing) : EMPTY_RELS;
		rels = rels.replace(/<Relationships\b([^>]*)\/>/u, '<Relationships$1></Relationships>');
		for (const id of [...page.usedMasters].sort((a, b) => a - b))
			rels = withRelationship(rels, `rIdMaster${id}`, MASTER_REL, `../masters/master${id}.xml`);
		await this.patch({ [part]: xml, [relsPart]: rels });

		const edits = page.takeEdits();
		const data = edits.filter((edit) => edit.type === 'set-shape-data');
		await this.edit(data, `${page.name}: shape data`);
		const other = edits.filter((edit) => edit.type !== 'set-shape-data');
		await this.edit(other, `${page.name}: hyperlinks`);
		await this.assignLayers(page);
	}

	private async assignLayers(page: PageBuilder): Promise<void> {
		const known = this.layerNames.get(page.pageId) ?? [];
		this.layerNames.set(page.pageId, known);
		const assign = (shapeIds: readonly number[], names: readonly string[], create = false) =>
			this.edit(
				[
					{
						type: 'assign-layers',
						pageId: page.pageId,
						shapeIds: shapeIds.map(String),
						layerIds: create ? [] : names.map((name) => String(known.indexOf(name))),
						...(create ? { newLayers: names } : {}),
					},
				],
				`${page.name}: layers ${names.join(', ')}`,
			);
		// Layers are created in the builder's order, so the indexes it wrote on groups hold.
		const missing = page.layerNames.filter((name) => !known.includes(name));
		if (missing.length > 0) {
			const first = [...page.layerSets.values()][0]?.[0];
			if (first === undefined)
				throw new Error(`${page.name}: a layer needs one shape that is not a group.`);
			await assign([first], missing, true);
			known.push(...missing);
		}
		for (const [key, shapeIds] of page.layerSets) await assign(shapeIds, key.split('|'));
		page.layerSets.clear();
	}

	/**
	 * Imports a table as a data recordset and links its rows to shapes, both through the library:
	 * each column becomes a linked Shape Data row of the shape. `links` pairs a shape ID with a
	 * one-based row number.
	 */
	async linkData(
		page: PageBuilder,
		spec: RecordsetSpec,
		links: readonly (readonly [shapeId: number, row: number])[],
	): Promise<void> {
		const recordsetId = String(this.recordsets.length);
		this.recordsets.push(spec);
		await this.edit(
			[
				{
					type: 'import-data-recordset',
					pageId: page.pageId,
					name: spec.name,
					columns: spec.columns.map((column) => ({ ...column })),
					rows: spec.rows.map((row) => [...row]),
					primaryKey: spec.primaryKey,
				},
			],
			`recordset ${spec.name}`,
		);
		await this.edit(
			[
				{
					type: 'link-data-rows',
					pageId: page.pageId,
					recordsetId,
					links: links.map(([shapeId, row]) => ({ shapeId: String(shapeId), rowId: String(row) })),
				},
			],
			`link ${spec.name}`,
		);
	}

	/** Replaces the plain text of the shapes whose fields were held back by text with fields. */
	async applyFields(page: PageBuilder): Promise<void> {
		const part = this.pagePart(page.pageId);
		let xml = await this.text(part);
		for (const [id, rich] of page.deferred) {
			const start = xml.search(new RegExp(`<Shape [^>]*ID="${id}"`, 'u'));
			const from = xml.indexOf('<Text>', start);
			const to = xml.indexOf('</Text>', from);
			const next = xml.indexOf('<Shape ', start + 1);
			if (start < 0 || from < 0 || to < 0 || (next >= 0 && next < from))
				throw new Error(`Shape ${id} has no text of its own to replace.`);
			xml = xml.slice(0, from) + textBody(rich) + xml.slice(to + '</Text>'.length);
		}
		page.deferred.clear();
		await this.patch({ [part]: xml });
	}

	/**
	 * Makes the library's recordset parts ones Visio works with and that do not depend on the
	 * clock: recordset IDs from 1 (Visio's object model refuses 0), a fixed refresh time, and the
	 * row parts rewritten with the row ID column.
	 */
	private async settleRecordsets(): Promise<void> {
		if (this.recordsets.length === 0) return;
		const part = 'visio/data/recordsets.xml';
		const xml = (await this.text(part))
			.replace(/NextID="\d+"/u, `NextID="${this.recordsets.length + 1}"`)
			.replaceAll(
				/<DataRecordSet ID="(\d+)"/gu,
				(_, id: string) => `<DataRecordSet ID="${Number(id) + 1}"`,
			)
			.replaceAll(/TimeRefreshed="[^"]*"/gu, `TimeRefreshed="${SAMPLE_DATE_ISO.slice(0, 19)}"`);
		const changes: Record<string, string> = { [part]: xml };
		const parts = await readPackage(this.bytes);
		this.recordsets.forEach((spec, index) => {
			const name = `visio/data/data${index + 1}.xml`;
			if (!parts.has(name)) throw new Error(`Expected the recordset part ${name}.`);
			changes[name] = rowsetXml(spec);
		});
		await this.patch(changes);
	}

	/** Adds the document properties and checks the rules every sample must keep. */
	async finish(): Promise<Uint8Array> {
		const { spec } = this;
		await this.settleRecordsets();
		await this.patch({
			'docProps/core.xml': corePropertiesXml(spec),
			'docProps/app.xml': APP_PROPERTIES_XML,
			'_rels/.rels': withPropertyRelationships(await this.text('_rels/.rels')),
			'[Content_Types].xml': withContentTypes(
				await this.text('[Content_Types].xml'),
				PROPERTY_CONTENT_TYPES,
			),
		});
		for (const [name, data] of await readPackage(this.bytes)) {
			const text = decoder.decode(data);
			const font = FORBIDDEN_FONTS.exec(text);
			if (font) throw new Error(`${name} names the font ${font[0]}.`);
			if (text.includes(EM_DASH)) throw new Error(`${name} contains an em-dash.`);
		}
		return this.bytes;
	}
}
