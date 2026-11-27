/**
 * Assembles a drawing: the core library creates the package and its pages and applies the
 * structural edits (containers, swimlanes, shape data, hyperlinks, layers, data links); this
 * module adds the parts the library has no writer for (styles, masters, drawn shapes, routes).
 */
import { SAMPLE_DATE } from '../../lib/types';
import { createVsdx } from '../../../../src/core/visio/create-document';
import { editVsdx } from '../../../../src/core/visio/edit';
import type { VisioEdit } from '../../../../src/core/visio/edit-commands';
import { patchPackage, readPackage } from '../../lib/package';
import {
	type DrawingSpec,
	assertSampleRules,
	documentPartChanges,
	pagePartNames,
	pageWithShapes,
} from './drawing-parts';
import { Stencil } from './masters';
import { PageBuilder } from './page';
import {
	APP_PROPERTIES_XML,
	EMPTY_RELS,
	PROPERTY_CONTENT_TYPES,
	corePropertiesXml,
	withContentTypes,
	withPropertyRelationships,
	withRelationship,
} from './parts';
import type { RecordsetSpec } from './recordset';
import { StyleBook } from './styles';
import { textBody } from './text';

const decoder = new TextDecoder();
const MASTER_REL = 'http://schemas.microsoft.com/visio/2010/relationships/master';
export type { DrawingSpec, PageSpec } from './drawing-parts';

export class Drawing {
	bytes: Uint8Array;
	readonly styles: StyleBook;
	readonly stencil: Stencil;
	/** Diagnostic codes the library reported while editing, for the sample's report. */
	readonly notes = new Set<string>();
	private readonly builders = new Map<string, PageBuilder>();
	private readonly pageParts = new Map<string, string>();
	private readonly layerNames = new Map<string, string[]>();
	/** Recordsets imported so far; the library numbers them from 1. */
	private recordsetCount = 0;

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
		for (const [id, part] of pagePartNames(pagesXml, rels)) this.pageParts.set(id, part);
	}

	private async writeDocumentParts(): Promise<void> {
		const read = (part: string) => this.text(part);
		await this.patch(await documentPartChanges(read, this.stencil, this.styles, this.spec));
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
			// The sample date is the time the edits stamp (a recordset's refresh time).
			const result = await editVsdx(this.bytes, [...edits], { now: SAMPLE_DATE });
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
		const xml = pageWithShapes(await this.text(part), shapes, connects);
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
		const recordsetId = String(++this.recordsetCount);
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

	/** Adds the document properties and checks the rules every sample must keep. */
	async finish(): Promise<Uint8Array> {
		const { spec } = this;
		await this.patch({
			'docProps/core.xml': corePropertiesXml(spec),
			'docProps/app.xml': APP_PROPERTIES_XML,
			'_rels/.rels': withPropertyRelationships(await this.text('_rels/.rels')),
			'[Content_Types].xml': withContentTypes(
				await this.text('[Content_Types].xml'),
				PROPERTY_CONTENT_TYPES,
			),
		});
		await assertSampleRules(this.bytes);
		return this.bytes;
	}
}
