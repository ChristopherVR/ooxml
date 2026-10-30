import { parseVsdx, type ParseVsdxOptions } from './parser';
import { VisioPackageError } from './package';
import type { VisioDiagnostic, VisioMaster } from './model';

/** The masters of an opened stencil file, as the Shapes window lists them. */
export interface VisioStencilFile {
	masters: VisioMaster[];
	diagnostics: VisioDiagnostic[];
}

/** True for the compound-file signature of legacy binary files (.vss, .vst, .vsd). */
const isCompoundFile = (bytes: Uint8Array) =>
	[0xd0, 0xcf, 0x11, 0xe0].every((value, index) => bytes[index] === value);

/**
 * Read a stencil package (.vssx, .vssm), a template (.vstx) or a drawing (.vsdx) for its masters: their
 * names, shapes, connection points, Shape Data and previews. Legacy binary stencils (.vss) are
 * refused, since the binary Visio formats belong to ole2. Macro-enabled packages (.vssm) read the same way and never run macros.
 */
export async function readVisioStencil(
	input: Uint8Array,
	options: Omit<ParseVsdxOptions, 'stencil'> = {},
): Promise<VisioStencilFile> {
	if (isCompoundFile(input))
		throw new VisioPackageError(
			'UNSUPPORTED_FORMAT',
			'Legacy binary stencils (.vss) are not supported; open a .vssx stencil.',
		);
	const document = await parseVsdx(input, { ...options, stencil: true });
	return { masters: document.masters ?? [], diagnostics: document.diagnostics };
}
