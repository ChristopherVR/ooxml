import { DEFAULTS, fail, type VisioPackageLimits } from './package-common.js';
import { visioXml } from './parts.js';
import { openEditablePackage, writeEditedPackage } from './edit-package.js';
import { replacePlainText, serializeEditedXml } from './edit-text.js';

export interface VisioTextEdit {
	type: 'replace-plain-text';
	pageId: string;
	shapeId: string;
	text: string;
}
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
/** Experimental source-backed plain-text transaction. No formulas are recalculated.
 * Untouched part payloads are byte-preserved; edited XML and ZIP representation are not.
 */
export async function editVsdx(
	input: Uint8Array | ArrayBuffer,
	edits: readonly VisioTextEdit[],
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
	if (!Array.isArray(edits) || edits.length > maxEdits)
		fail('LIMIT_EDITS', 'Edit command count exceeds limit.');
	let textLength = 0;
	// Copy commands before the first await: caller mutation cannot change the transaction.
	const commands = Array.from(edits, (edit) => {
		if (
			!edit ||
			edit.type !== 'replace-plain-text' ||
			typeof edit.pageId !== 'string' ||
			!edit.pageId ||
			edit.pageId.length > 256 ||
			typeof edit.shapeId !== 'string' ||
			!edit.shapeId ||
			edit.shapeId.length > 256 ||
			typeof edit.text !== 'string'
		)
			fail('INVALID_EDIT', 'Invalid plain-text edit command.');
		textLength += edit.text.length;
		if (textLength > maxText) fail('LIMIT_EDIT_TEXT', 'Replacement text exceeds aggregate limit.');
		if (
			/[\u0000-\u0008\u000b\u000c\u000d\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u.test(
				edit.text,
			)
		)
			fail('INVALID_EDIT_TEXT', 'Replacement text contains invalid XML characters.');
		return { pageId: edit.pageId, shapeId: edit.shapeId, text: edit.text };
	});
	const source = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (source.length > limits.maxInputBytes) fail('LIMIT_INPUT', 'ZIP input exceeds limit.');
	const original = new Uint8Array(source);
	const { pkg, parts, pages } = await openEditablePackage(original, limits, check);
	const dirty = new Map<string, Element>();
	for (const command of commands) {
		check();
		const path = pages.get(command.pageId);
		if (!path) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
		let root = dirty.get(path);
		if (!root) {
			const sourceRoot = await visioXml(pkg, path, 'PageContents');
			root = (sourceRoot.ownerDocument!.cloneNode(true) as Document).documentElement;
		}
		if (replacePlainText(root, command.shapeId, command.text, check)) dirty.set(path, root);
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
		diagnostics: dirty.size
			? [
					{
						code: 'edit-caches-not-recalculated',
						message:
							'Text was replaced without recalculating formulas or dependent caches. Native Visio reopen and rendering compatibility are unverified.',
					},
				]
			: [],
	};
}
