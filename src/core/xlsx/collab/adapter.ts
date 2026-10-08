// The whole-model `DocumentAdapter<Workbook>` for `bindDocument`: `write` reconciles the full
// workbook into the shared maps (only differences are written), `read` materialises a workbook.
// Products with an edit session should prefer `bindWorkbookSession`, which writes per edit.
import type * as Y from 'yjs';
import type { DocumentAdapter } from '../../collab/binding';
import type { Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { applyShared } from './apply';
import { emptyDirty } from './dirty';
import { sharedTypes, sharedWorkbook } from './schema';
import { SheetKeys } from './sheets';
import { isCellStyle } from './styles';
import { writeWorkbook } from './write';

export interface XlsxAdapterOptions {
	/**
	 * The workbook `read` starts from (cloned), so what the shared schema does not carry (charts,
	 * comments, conditional formats, ...) survives; default a new empty workbook.
	 */
	base?: () => Workbook;
}

/** Whether the shared document holds no workbook yet. */
export const isSharedWorkbookEmpty = (doc: Y.Doc): boolean => sharedWorkbook(doc).sheets.size === 0;

/** Materialises the shared workbook (onto a clone of `base` when given). */
export function readSharedWorkbook(doc: Y.Doc, base?: Workbook): Workbook {
	const shared = sharedWorkbook(doc);
	const workbook = base ? structuredClone(base) : createWorkbook();
	if (!base) {
		const style = shared.styles.get(String(shared.meta.get('defaultStyle')));
		if (isCellStyle(style)) workbook.styles = [structuredClone(style)];
	}
	applyShared(shared, new SheetKeys(doc.clientID), workbook, emptyDirty(true));
	return workbook;
}

/** Writes the whole workbook into the shared document (differences only). */
export function writeSharedWorkbook(doc: Y.Doc, workbook: Workbook, origin?: unknown): void {
	doc.transact(() => {
		writeWorkbook(sharedWorkbook(doc), new SheetKeys(doc.clientID), workbook, { kind: 'all' });
	}, origin);
}

export function xlsxDocumentAdapter(options: XlsxAdapterOptions = {}): DocumentAdapter<Workbook> {
	return {
		isEmpty: isSharedWorkbookEmpty,
		read: (doc) => readSharedWorkbook(doc, options.base?.()),
		write: (doc, model, origin) => writeSharedWorkbook(doc, model, origin),
		observe: (doc, onChange) => {
			const types = sharedTypes(sharedWorkbook(doc));
			let origin: unknown;
			let changed = false;
			const mark = (_events: unknown, transaction: Y.Transaction): void => {
				changed = true;
				origin = transaction.origin;
			};
			const flush = (): void => {
				if (!changed) return;
				changed = false;
				onChange(origin);
			};
			for (const type of types) type.observeDeep(mark);
			doc.on('afterTransaction', flush);
			return () => {
				for (const type of types) type.unobserveDeep(mark);
				doc.off('afterTransaction', flush);
			};
		},
	};
}
