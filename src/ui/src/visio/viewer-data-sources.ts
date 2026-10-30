import {
	visioRefreshReport,
	type VisioDataRecordset,
	type VisioEdit,
	type VisioPage,
	type VisioRefreshReport,
} from 'ooxml-core/visio';
import type { ViewerController } from './controller';
import type { InsertDialog } from './viewer-insert-dialog';
import {
	createImportDialog,
	DataFilePicker,
	readTable,
	readWorkbook,
	type DataFileHandle,
	type DataSourceOptions,
	type PickedFile,
} from './viewer-data-import';
import type { ViewerDataLinks } from './viewer-data-links';

type Run = (action: () => Promise<void>, success: string) => void;

/**
 * Quick Import, Custom Import and Refresh All. A recordset remembers how it was read and, when
 * the browser offers one, its File System Access handle; without a handle Refresh asks for the
 * file again, because a page cannot re-read a local path by itself.
 */
export class ViewerDataSources {
	readonly picker: DataFilePicker;
	readonly importer: InsertDialog;
	#sources = new Map<
		string,
		{ name: string; handle?: DataFileHandle; options: DataSourceOptions }
	>();
	#pendingImport:
		| { picked: PickedFile; workbook: Awaited<ReturnType<typeof readWorkbook>> }
		| undefined;
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
		private readonly run: Run,
		private readonly links: ViewerDataLinks,
	) {
		this.picker = new DataFilePicker(root);
		this.importer = createImportDialog(root, (button) => void this.#customImport(button));
	}
	#page(): VisioPage | undefined {
		const state = this.controller.state;
		return state.document?.pages[state.pageIndex];
	}
	async import(custom: boolean): Promise<void> {
		const picked = await this.picker.pick();
		if (!picked) return;
		let workbook;
		try {
			workbook = await readWorkbook(picked.file);
		} catch (error) {
			return this.announce(
				`${picked.file.name} cannot be imported: ${error instanceof Error ? error.message : 'unreadable file'}`,
			);
		}
		const name = picked.file.name.slice(0, 255);
		if (!custom)
			return this.#importTable(picked, workbook, { sheet: 0, range: '', header: true }, name);
		this.#pendingImport = { picked, workbook };
		(this.importer.fields.get('name') as HTMLInputElement).value = name;
		this.importer.choices(
			'sheet',
			workbook.sheets.map((_, index) => String(index)),
			'0',
			{ labels: workbook.sheets.map((sheet) => sheet.name), none: false },
		);
		(this.importer.fields.get('range') as HTMLInputElement).value = '';
		(this.importer.fields.get('key') as HTMLInputElement).value = '';
		(this.importer.fields.get('header') as HTMLInputElement).checked = true;
		this.importer.error.textContent = '';
		this.importer.dialog.show();
		this.importer.fields.get('name')!.focus();
	}
	async #customImport(button: string): Promise<void> {
		const pending = this.#pendingImport;
		if (button === 'Cancel' || !pending) return this.importer.dialog.close();
		const name = this.importer.value('name').trim();
		if (!name) return void (this.importer.error.textContent = 'Type a name for the recordset.');
		const options = {
			sheet: Number(this.importer.value('sheet')) || 0,
			range: this.importer.value('range'),
			header: this.importer.checked('header'),
		};
		const key = this.importer.value('key').trim();
		try {
			const table = readTable(pending.workbook, options);
			if (key && !table.columns.some((column) => column.name === key || column.label === key))
				throw new Error(`The data has no column named ${key}.`);
		} catch (error) {
			return void (this.importer.error.textContent =
				error instanceof Error ? error.message : String(error));
		}
		this.importer.dialog.close();
		this.#pendingImport = undefined;
		await this.#importTable(pending.picked, pending.workbook, options, name, key);
	}
	async #importTable(
		picked: PickedFile,
		workbook: Awaited<ReturnType<typeof readWorkbook>>,
		options: DataSourceOptions,
		name: string,
		key = '',
	): Promise<void> {
		const page = this.#page();
		if (!page) return;
		let table;
		try {
			table = readTable(workbook, options);
		} catch (error) {
			return this.announce(
				`${picked.file.name} cannot be imported: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
		const column = key
			? table.columns.find((item) => item.name === key || item.label === key)
			: undefined;
		const csv = /\.(csv|txt)$/i.test(picked.file.name);
		const range = options.range.trim().toUpperCase();
		const before = new Set(
			(this.controller.state.document?.dataRecordsets ?? []).map((set) => set.id),
		);
		this.run(async () => {
			await this.controller.applyEdits([
				{
					type: 'import-data-recordset',
					pageId: page.id,
					name,
					...table,
					// Saved as the recordset's data connection, command and primary key.
					source: {
						kind: csv ? 'csv' : 'workbook',
						fileName: picked.file.name,
						...(csv ? {} : { sheet: workbook.sheets[options.sheet]?.name ?? 'Sheet1' }),
						...(range ? { range } : {}),
					},
					...(column ? { primaryKey: column.name } : {}),
				},
			]);
			const added = (this.controller.state.document?.dataRecordsets ?? []).find(
				(set) => !before.has(set.id),
			);
			if (!added) return;
			this.#sources.set(added.id, {
				name: picked.file.name,
				options,
				...(picked.handle ? { handle: picked.handle } : {}),
			});
			this.links.show(added.id);
		}, `Imported ${table.rows.length} rows from ${picked.file.name}. Drag a row onto a shape to link it.`);
	}
	async refresh(): Promise<void> {
		const sets = this.controller.state.document?.dataRecordsets ?? [];
		const page = this.#page();
		if (!sets.length || !page) return this.announce('There is no external data to refresh.');
		const edits: VisioEdit[] = [];
		const reports: { set: VisioDataRecordset; report: VisioRefreshReport }[] = [];
		const missing: VisioDataRecordset[] = [];
		for (const set of sets) {
			const source = this.#sources.get(set.id);
			if (!source?.handle) {
				missing.push(set);
				continue;
			}
			try {
				const table = readTable(await readWorkbook(await source.handle.getFile()), source.options);
				const report = visioRefreshReport(set, table.columns, table.rows);
				if (report.keyMissing)
					return this.announce(
						`${set.name} cannot be refreshed: the data has no ${report.key} column, its unique identifier.`,
					);
				reports.push({ set, report });
				edits.push({
					type: 'refresh-data-recordset',
					pageId: page.id,
					recordsetId: set.id,
					...table,
				});
			} catch (error) {
				return this.announce(
					`${set.name} cannot be refreshed: ${error instanceof Error ? error.message : String(error)}`,
				);
			}
		}
		const active = missing.find((set) => set.id === this.links.activeId) ?? missing[0];
		if (!edits.length && active) {
			// Without a saved file handle the browser cannot read the original path again.
			this.announce(
				`Choose ${active.name} again to refresh it; this browser cannot re-read the original file.`,
			);
			const picked = await this.picker.pick();
			if (!picked) return;
			try {
				const options = this.#sources.get(active.id)?.options ?? {
					sheet: 0,
					range: '',
					header: true,
				};
				const table = readTable(await readWorkbook(picked.file), options);
				const report = visioRefreshReport(active, table.columns, table.rows);
				if (report.keyMissing)
					return this.announce(
						`${active.name} cannot be refreshed: the data has no ${report.key} column, its unique identifier.`,
					);
				reports.push({ set: active, report });
				this.#sources.set(active.id, {
					name: picked.file.name,
					options,
					...(picked.handle ? { handle: picked.handle } : {}),
				});
				edits.push({
					type: 'refresh-data-recordset',
					pageId: page.id,
					recordsetId: active.id,
					...table,
				});
			} catch (error) {
				return this.announce(
					`${picked.file.name} cannot be read: ${error instanceof Error ? error.message : String(error)}`,
				);
			}
		}
		this.run(
			() => this.controller.applyEdits(edits),
			`Refreshed ${edits.length === 1 ? '1 recordset' : `${edits.length} recordsets`}${missing.length && edits.length < sets.length ? '; choose the other files again to refresh them' : ''}.${reports.map(refreshNote).join('')}`,
		);
	}
}

/** What a refresh did to linked shapes: how rows were matched and which links were lost. */
function refreshNote({ set, report }: { set: VisioDataRecordset; report: VisioRefreshReport }) {
	const lost = report.unmatched.length;
	const by = report.key ? `by ${report.key}` : 'by row order';
	const parts = [`${set.name}: matched ${report.matched} linked shapes ${by}`];
	if (lost)
		parts.push(
			`${lost} linked ${lost === 1 ? 'shape has' : 'shapes have'} no matching row and ${lost === 1 ? 'was' : 'were'} unlinked (shape ${report.unmatched
				.slice(0, 5)
				.map((link) => link.shapeId)
				.join(', ')}${lost > 5 ? ', ...' : ''})`,
		);
	if (report.added)
		parts.push(`${report.added} new ${report.added === 1 ? 'row' : 'rows'} to link`);
	return ` ${parts.join('; ')}.`;
}
