import { NS } from '../xml/index.js';
import { inspectVisioEmfAdmission } from './emf-admission.js';
import type { VisioPackage } from './package.js';
import { children, type Report } from './sheet.js';

/** Inspection budgets across unique internal media parts, independently of ZIP/XML limits. */
export const VISIO_METAFILE_INSPECTION_LIMITS = Object.freeze({
	maxAssets: 32,
	maxTotalBytes: 32 * 1024 * 1024,
	maxTotalRecords: 100_000,
});
type Note = { code: string; message: string };
interface State {
	bytes: number;
	records: number;
	results: Map<string, Promise<readonly Note[]>>;
	tail: Promise<void>;
}
const states = new WeakMap<VisioPackage, State>();
/** Detailed fail-closed compatibility inspection only. Never invokes a converter or returns drawable media. */
export async function inspectEmbeddedVisioMetafile(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
): Promise<void> {
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel) {
		report('invalid-image-relationship', 'Metafile has no unambiguous relationship reference.');
		return;
	}
	if (rel.mode !== 'Internal') {
		report('external-image', 'External metafile relationships are not loaded.');
		return;
	}
	if (rel.type !== `${NS.r}/image`) {
		report('invalid-image-relationship', 'Metafile does not reference an image relationship.');
		return;
	}
	let state = states.get(pkg);
	if (!state) {
		state = { bytes: 0, records: 0, results: new Map(), tail: Promise.resolve() };
		states.set(pkg, state);
	}
	let pending = state.results.get(rel.target);
	if (!pending) {
		if (state.results.size >= VISIO_METAFILE_INSPECTION_LIMITS.maxAssets) {
			report(
				'emf-document-limit',
				'Additional metafile parts were not inspected because the document asset limit was reached.',
				{ part: rel.target },
			);
			return;
		}
		const current = state;
		pending = current.tail.then(() => inspectPart(pkg, rel.target, current));
		// Reserve a unique slot before yielding; duplicate calls share the same promise.
		current.results.set(rel.target, pending);
		current.tail = pending.then(
			() => undefined,
			() => undefined,
		);
	}
	for (const note of await pending) report(note.code, note.message, { part: rel.target });
}

async function inspectPart(
	pkg: VisioPackage,
	part: string,
	state: State,
): Promise<readonly Note[]> {
	let notes: readonly Note[];
	const size = pkg.getPartByteLength(part);
	const remaining = VISIO_METAFILE_INSPECTION_LIMITS.maxTotalRecords - state.records;
	if (size > 4 * 1024 * 1024) {
		notes = [
			{
				code: 'emf-input-limit',
				message: 'Metafile bytes exceed the per-asset limit; the part was not inflated.',
			},
		];
	} else if (
		size > VISIO_METAFILE_INSPECTION_LIMITS.maxTotalBytes - state.bytes ||
		remaining <= 0
	) {
		notes = [
			{
				code: 'emf-document-limit',
				message:
					'Metafile inspection exceeded the aggregate byte or record limit; the part was not inflated.',
			},
		];
	} else {
		const bytes = await pkg.readBytes(part);
		state.bytes += bytes.byteLength;
		const result = inspectVisioEmfAdmission(bytes, { maxRecords: Math.min(20_000, remaining) });
		state.records += result.metrics.records;
		notes = result.diagnostics.map((note) => ({
			code: `emf-${note.code}`,
			message: `${note.message} Record ${note.recordType ?? 'header'} at byte ${note.offset}.`,
		}));
		if (result.omittedDiagnostics)
			notes = [
				...notes,
				{
					code: 'emf-diagnostics-truncated',
					message: `${result.omittedDiagnostics} additional metafile inspection diagnostics were omitted.`,
				},
			];
		if (result.status === 'admitted')
			notes = [
				...notes,
				{
					code: 'emf-rendering-disabled',
					message:
						'This metafile passes the narrow structural preflight, but vector conversion and fidelity are not enabled.',
				},
			];
	}
	return notes;
}
