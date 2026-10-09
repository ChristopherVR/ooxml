import { DEFAULTS, fail, type VisioPackageLimits } from './package-common';
import { visioXml, related } from './parts';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import { replaceScopedPlainText } from './edit-text-scope';
import { replaceScopedTextRanges } from './edit-text-ranges';
import {
	snapshotVisioEdits,
	isVisioPageEdit,
	type VisioEdit,
	type VisioGeometryEdit,
} from './edit-commands';
import { applyGeometryEdit } from './edit-geometry';
import { assertDuplicateScope } from './edit-duplicate-scope';
import { assertGeometryPackageScope } from './edit-scope';
import { emptyMasterMoveProof } from './edit-master-move';
import { editVsdxPages } from './edit-pages';
import { applyFormattingEdit } from './edit-formatting';
import { isVisioFormatEdit } from './edit-formatting-commands';
import { reorderVisioShape, assertShapeOrderPackageScope } from './edit-shape-order';
import { duplicateVisioShapes } from './edit-duplicate';
import { pasteVisioShapes } from './edit-paste';
import { deleteVisioShapes, type VisioShapeDelete } from './edit-delete';
import { changeVisioShape } from './edit-change-shape';
export type {
	VisioEdit,
	VisioTextEdit,
	VisioTextRange,
	VisioTextRangesEdit,
	VisioGeometryEdit,
	VisioPageInsert,
	VisioPageReorder,
	VisioPageRename,
	VisioPageDelete,
	VisioPageSizeEdit,
	VisioPageEdit,
	VisioFormatEdit,
	VisioTextFormatEdit,
	VisioShapeFormatEdit,
	VisioShapeOrderEdit,
	VisioDuplicateShapesEdit,
	VisioPasteShapesEdit,
	VisioResizeAnchor,
	VisioChangeShapeEdit,
	VisioChangeShapeTarget,
} from './edit-commands';

export interface EditVsdxOptions {
	limits?: Partial<VisioPackageLimits>;
	maxEdits?: number;
	maxTextCharacters?: number;
	maxOutputBytes?: number;
}
export interface EditVsdxResult {
	bytes: Uint8Array;
	changedParts: readonly string[];
	diagnostics: readonly { code: string; message: string }[];
}
function positive(value: number): number {
	if (!Number.isSafeInteger(value) || value <= 0)
		fail('INVALID_LIMITS', 'Edit limits must be positive safe integers.');
	return value;
}
/** Experimental source-backed atomic text and conservative geometry transaction.
 * Untouched part payloads are byte-preserved; edited XML and ZIP representation are not.
 */
export async function editVsdx(
	input: Uint8Array | ArrayBuffer,
	edits: readonly VisioEdit[],
	options: EditVsdxOptions = {},
): Promise<EditVsdxResult> {
	const limits = { ...DEFAULTS, ...options.limits };
	for (const value of Object.values(limits)) positive(value);
	const maxEdits = positive(options.maxEdits ?? 1000);
	const maxText = positive(options.maxTextCharacters ?? 1_000_000);
	const maxOutput = positive(options.maxOutputBytes ?? limits.maxInputBytes);
	const deadline = Date.now() + limits.maxRuntimeMs;
	const check = () => {
		if (Date.now() >= deadline) fail('LIMIT_RUNTIME', 'Visio edit deadline exceeded.');
	};
	// Copy commands before the first await: caller mutation cannot change the transaction.
	const allCommands = snapshotVisioEdits(edits, maxEdits, maxText);
	const commands = allCommands.filter((command) => !isVisioPageEdit(command));
	const source = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (source.length > limits.maxInputBytes) fail('LIMIT_INPUT', 'ZIP input exceeds limit.');
	const original = new Uint8Array(source);
	const { pkg, parts, pages } = await openEditablePackage(original, limits, check);
	if (commands.length !== allCommands.length) {
		if (commands.length)
			fail(
				'EDIT_MIXED_PAGE_TRANSACTION',
				'Page edits and shape edits require separate transactions.',
			);
		return editVsdxPages(
			original,
			pkg,
			parts,
			allCommands.filter(isVisioPageEdit),
			limits,
			maxOutput,
			deadline,
			check,
		);
	}
	const dirty = new Map<string, Element>();
	const roots = new Map<string, Element>();
	const geometryCommands = commands.filter(
		(command): command is VisioGeometryEdit =>
			command.type !== 'replace-plain-text' &&
			command.type !== 'replace-text-ranges' &&
			command.type !== 'reorder-shape' &&
			command.type !== 'duplicate-shapes' &&
			command.type !== 'paste-shapes' &&
			command.type !== 'change-shape' &&
			!isVisioFormatEdit(command),
	);
	let document: Element | undefined;
	let masterMovePins = emptyMasterMoveProof();
	if (
		geometryCommands.length ||
		commands.some(
			(command) =>
				command.type === 'replace-plain-text' ||
				command.type === 'replace-text-ranges' ||
				isVisioFormatEdit(command) ||
				command.type === 'reorder-shape' ||
				command.type === 'duplicate-shapes' ||
				command.type === 'paste-shapes' ||
				command.type === 'change-shape',
		)
	) {
		// All pages are indexed before editing: dependencies are never inferred from only the target shape.
		for (const [pageId, path] of pages) {
			const sourceRoot = await visioXml(pkg, path, 'PageContents');
			roots.set(pageId, (sourceRoot.ownerDocument!.cloneNode(true) as Document).documentElement);
		}
		for (const [pageId, root] of roots) {
			const ids = new Set(
				geometryCommands
					.filter((command) => command.type === 'create-text-box' && command.pageId === pageId)
					.map((command) => command.shapeId),
			);
			if (ids.size) await assertDuplicateScope(pkg, new Set(pages.values()), root, ids, check);
		}
		if (geometryCommands.length)
			masterMovePins = await assertGeometryPackageScope(
				pkg,
				new Set(pages.values()),
				geometryCommands,
				check,
				roots,
			);
		const path = await related(pkg, '', 'document');
		document = await visioXml(pkg, path!, 'VisioDocument');
	}
	let textChanged = false;
	if (commands.some((command) => command.type === 'reorder-shape'))
		await assertShapeOrderPackageScope(pkg, check);
	let formatChanged = false;
	let orderChanged = false;
	let duplicateChanged = false;
	let pasteChanged = false;
	let outlineChanged = false;
	const deletions = commands.filter(
		(command): command is VisioShapeDelete => command.type === 'delete-shape',
	);
	const deleteOnly = deletions.length > 0 && deletions.length === commands.length;
	if (deleteOnly)
		for (const pageId of deleteVisioShapes(roots, document!, deletions, check))
			dirty.set(pages.get(pageId)!, roots.get(pageId)!);
	for (const command of deleteOnly ? [] : commands) {
		check();
		const path = pages.get(command.pageId);
		if (!path) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
		let root = roots.get(command.pageId) ?? dirty.get(path);
		if (!root) {
			const sourceRoot = await visioXml(pkg, path, 'PageContents');
			root = (sourceRoot.ownerDocument!.cloneNode(true) as Document).documentElement;
			roots.set(command.pageId, root);
		}
		if (command.type === 'replace-plain-text' || command.type === 'replace-text-ranges') {
			if (
				await (command.type === 'replace-plain-text'
					? replaceScopedPlainText(pkg, new Set(pages.values()), roots, document!, command, check)
					: replaceScopedTextRanges(pkg, new Set(pages.values()), roots, document!, command, check))
			) {
				dirty.set(path, root);
				textChanged = true;
			}
		} else if (isVisioFormatEdit(command)) {
			if (
				await applyFormattingEdit(pkg, new Set(pages.values()), roots, document!, command, check)
			) {
				dirty.set(path, root);
				formatChanged = true;
			}
		} else if (command.type === 'reorder-shape') {
			if (reorderVisioShape(root, document!, command, check)) {
				dirty.set(path, root);
				orderChanged = true;
			}
		} else if (command.type === 'change-shape') {
			if (changeVisioShape(roots, document!, command, check)) {
				dirty.set(path, root);
				outlineChanged = true;
			}
		} else if (command.type === 'paste-shapes') {
			for (const pageId of await pasteVisioShapes(
				pkg,
				new Set(pages.values()),
				roots,
				document!,
				command,
				check,
			))
				dirty.set(pages.get(pageId)!, roots.get(pageId)!);
			pasteChanged = true;
		} else if (command.type === 'duplicate-shapes') {
			for (const pageId of await duplicateVisioShapes(
				pkg,
				new Set(pages.values()),
				roots,
				document!,
				command,
				check,
			))
				dirty.set(pages.get(pageId)!, roots.get(pageId)!);
			duplicateChanged = true;
		} else {
			for (const pageId of applyGeometryEdit(roots, document!, command, check, masterMovePins))
				dirty.set(pages.get(pageId)!, roots.get(pageId)!);
		}
	}
	let totalBytes = [...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0);
	let nodes = 0;
	for (const [path, root] of dirty) {
		const serialized = serializeEditedXml(root, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		totalBytes += serialized.bytes.length - parts.get(path)!.length;
		if (totalBytes > limits.maxTotalBytes)
			fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
		parts.set(path, serialized.bytes);
	}
	check();
	if (!dirty.size && original.length > maxOutput)
		fail('LIMIT_EDIT_OUTPUT', 'Saved package exceeds output limit.');
	const bytes = dirty.size ? await writeEditedPackage(parts, maxOutput, deadline, check) : original;
	if (dirty.size) {
		// Re-admit the saved package and count metadata plus changed page XML together.
		// A fresh reader gets only the remaining transaction time, never a new budget.
		check();
		const verified = await openEditablePackage(
			bytes,
			{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
			check,
		);
		for (const path of dirty.keys()) await visioXml(verified.pkg, path, 'PageContents');
		check();
	}
	return {
		bytes,
		changedParts: [...dirty.keys()],
		diagnostics: [
			...(pasteChanged
				? [
						{
							code: 'edit-paste-experimental',
							message:
								'Captured shape XML was pasted with compatible resources and page context. Clipboard placement and cross-document fidelity are limited to supported cases.',
						},
					]
				: []),
			...(duplicateChanged
				? [
						{
							code: 'edit-duplicate-experimental',
							message:
								'Local shape XML was copied and supported pin-dependent caches were recalculated. Native fidelity is limited to tested cases.',
						},
					]
				: []),
			...(dirty.size && textChanged
				? [
						{
							code: 'edit-caches-not-recalculated',
							message:
								'Plain text changed; affected or unknown formula dependencies were refused. Untouched caches were preserved. Native fidelity is limited to tested cases.',
						},
					]
				: []),
			...(dirty.size && deleteOnly
				? [
						{
							code: 'edit-delete-experimental',
							message:
								'Reference-closed local shapes were deleted. Retained formulas and caches were preserved; glue healing and inherited shape deletion remain unsupported.',
						},
					]
				: []),
			...(dirty.size && geometryCommands.length && !deleteOnly
				? [
						{
							code: 'edit-geometry-experimental',
							message:
								'Supported affected geometry caches were recalculated. Native Visio reopen and rendering fidelity remain unverified.',
						},
					]
				: []),
			...(formatChanged
				? [
						{
							code: 'edit-formatting-experimental',
							message:
								'Text and solid shape formatting was changed without recalculating text layout. Native Visio reopen and rendering fidelity remain unverified.',
						},
					]
				: []),
			...(outlineChanged
				? [
						{
							code: 'edit-change-shape-experimental',
							message:
								'Local Geometry sections were replaced with a Basic Shapes outline. Native Visio reopen and rendering fidelity remain unverified.',
						},
					]
				: []),
			...(orderChanged
				? [
						{
							code: 'edit-shape-order-experimental',
							message:
								'Local sibling shape order was changed. Native Visio reopen and rendering fidelity remain unverified.',
						},
					]
				: []),
		],
	};
}
