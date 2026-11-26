import { DEFAULTS, fail, type VisioPackageLimits } from './package-common';
import { prepareVisioPackageForEdit, STENCIL_MACRO_MAIN } from './package-kind';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { related } from './parts';
import { serializeEditedXml } from './edit-text';

export interface VisioStencilEditOptions {
	limits?: Partial<VisioPackageLimits>;
}

export function stencilSession(options: VisioStencilEditOptions) {
	const limits = { ...DEFAULTS, ...options.limits };
	for (const value of Object.values(limits))
		if (!Number.isSafeInteger(value) || value < 1)
			fail('INVALID_LIMITS', 'Visio stencil limits require positive safe integers.');
	const deadline = Date.now() + limits.maxRuntimeMs;
	const check = () => {
		if (Date.now() >= deadline) fail('LIMIT_RUNTIME', 'Visio stencil deadline exceeded.');
	};
	return { limits, deadline, check };
}
export type StencilSession = ReturnType<typeof stencilSession>;

export async function openStencil(source: Uint8Array, session: StencilSession) {
	const { limits, check } = session;
	const prepared = await prepareVisioPackageForEdit(source, limits);
	const workspace = await openEditablePackage(prepared.bytes, limits, check, true);
	const documentPart = await related(workspace.pkg, '', 'document');
	if (!documentPart) fail('INVALID_DOCUMENT', 'Missing Visio document part.');
	const mastersPart = await related(workspace.pkg, documentPart, 'masters', false);
	const extension: 'vssx' | 'vssm' = workspace.mainType === STENCIL_MACRO_MAIN ? 'vssm' : 'vssx';
	return {
		...workspace,
		mastersPart,
		extension,
		source: prepared.bytes,
		signatureRemoved: prepared.signatureRemoved,
	};
}

export async function saveStencilParts(
	parts: Map<string, Uint8Array>,
	dirty: ReadonlyMap<string, Element>,
	{ limits, check, deadline }: StencilSession,
): Promise<Uint8Array> {
	let nodes = 0;
	for (const [path, xml] of dirty) {
		const result = serializeEditedXml(xml, limits, check);
		nodes += result.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		parts.set(path, result.bytes);
	}
	let bytes = 0;
	for (const part of parts.values()) bytes += part.length;
	if (bytes > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
	if (parts.size > limits.maxEntries) fail('LIMIT_ENTRIES', 'Edited package entry limit exceeded.');
	return writeEditedPackage(parts, limits.maxInputBytes, deadline, check);
}
