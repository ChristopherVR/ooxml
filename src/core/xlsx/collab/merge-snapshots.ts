import * as Y from 'yjs';
import { mergeSnapshotMaps, mergeSnapshotValue, snapshotEqual } from '../../collab/merge-values';
import { mergeOpcSnapshots, type SnapshotMergeResult } from '../../opc/merge-snapshots';
import { loadXlsx } from '../read/load';
import { readZipParts, SourceIndex } from '../read/package';
import { saveXlsx } from '../write/save';
import { readSharedWorkbook, writeSharedWorkbook } from './adapter';
import {
	applyWorkbookResidual,
	stripFormulaCaches,
	unsupportedSourceChanges,
	workbookResidual,
} from './merge-residual';

/**
 * Reconciles saved XLSX files by logical sheet/cell/style identity using a common Yjs baseline.
 * This is a three-way merge of snapshots, not recovery of missing collaboration operations.
 * Arrays outside the stable mapping and changes to the same field remain unresolved.
 */
export async function mergeWorkbookSnapshots(
	base: Uint8Array,
	local: Uint8Array,
	remote: Uint8Array,
): Promise<SnapshotMergeResult> {
	const docs: Y.Doc[] = [];
	try {
		const books = await Promise.all([base, local, remote].map((bytes) => loadXlsx(bytes)));
		const baseline = books[0]!,
			left = books[1]!,
			right = books[2]!;
		const issues = unsupportedSourceChanges(books);
		if (
			!snapshotEqual(baseline.styles[0], left.styles[0]) ||
			!snapshotEqual(baseline.styles[0], right.styles[0])
		)
			issues.push('Changing the default workbook style cannot be reconciled from saved snapshots.');
		const seed = new Y.Doc();
		docs.push(seed);
		writeSharedWorkbook(seed, baseline);
		const localDoc = new Y.Doc(),
			remoteDoc = new Y.Doc();
		docs.push(localDoc, remoteDoc);
		for (const doc of [localDoc, remoteDoc]) Y.applyUpdate(doc, Y.encodeStateAsUpdate(seed));
		writeSharedWorkbook(localDoc, left);
		writeSharedWorkbook(remoteDoc, right);
		const residue = mergeSnapshotValue(
			workbookResidual(baseline, seed),
			workbookResidual(left, localDoc),
			workbookResidual(right, remoteDoc),
			'workbook',
			issues,
		) as ReturnType<typeof workbookResidual>;
		const formulas = [seed, localDoc, remoteDoc].map(stripFormulaCaches).some(Boolean);
		const merged = mergeSnapshotMaps(seed, localDoc, remoteDoc, [...seed.share.keys()]);
		docs.push(merged.doc);
		issues.push(...merged.issues);
		if (issues.length) return { status: 'unresolved', issues };
		const paths = new Set(
			books.flatMap((book) => {
				const source = book.source && new SourceIndex(book.source.parts);
				return [...(source?.parts.keys() ?? [])].filter((path) => {
					const type = source?.contentType(path) ?? '';
					return /spreadsheetml\.(?:sheet\.main|worksheet|styles|sharedStrings|comments|table|drawing|chart)\+xml$/u.test(
						type,
					);
				});
			}),
		);
		const opaque = await mergeOpcSnapshots(base, local, remote, {
			mergePart: (part) =>
				paths.has(part.path)
					? { status: 'merged', bytes: part.remote }
					: {
							status: 'unresolved',
							issues: [`Unmapped package part changed concurrently: ${part.path}.`],
						},
		});
		if (opaque.status === 'unresolved') return opaque;
		const workbook = readSharedWorkbook(merged.doc, baseline);
		applyWorkbookResidual(workbook, merged.doc, residue);
		for (const key of ['modified', 'lastModifiedBy'] as const) {
			const value = right.properties[key];
			if (value === undefined) delete workbook.properties[key];
			else workbook.properties[key] = value;
		}
		workbook.source = { parts: await readZipParts(opaque.bytes) };
		if (formulas) {
			workbook.valuesStale = true;
			workbook.fullCalcOnLoad = true;
		}
		return { status: 'merged', bytes: await saveXlsx(workbook) };
	} catch {
		return {
			status: 'unresolved',
			issues: ['The XLSX snapshots could not be safely parsed or combined.'],
		};
	} finally {
		for (const doc of docs) doc.destroy();
	}
}
