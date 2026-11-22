import { openZip } from './zip-limits';

const OFFICE_PACKAGE_FORMATS = [
	{
		mainPart: 'word/document.xml',
		format: {
			extension: 'docx',
			type: 'word',
			progId: 'Word.Document.12',
			mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		},
	},
	{
		mainPart: 'xl/workbook.xml',
		format: {
			extension: 'xlsx',
			type: 'excel',
			progId: 'Excel.Sheet.12',
			mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		},
	},
	{
		mainPart: 'ppt/presentation.xml',
		format: {
			extension: 'pptx',
			type: 'powerpoint',
			progId: 'PowerPoint.Show.12',
			mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		},
	},
] as const;

export type EmbeddableOfficePackage = (typeof OFFICE_PACKAGE_FORMATS)[number]['format'];

export const MAX_LOCAL_OBJECT_BYTES = 32 * 1024 * 1024;

function findOfficePackageFormat(extension: string) {
	const normalized = extension.toLowerCase();
	return OFFICE_PACKAGE_FORMATS.find(({ format }) => format.extension === normalized);
}

export function embeddableOfficePackageFormat(
	extension: string,
): EmbeddableOfficePackage | undefined {
	return findOfficePackageFormat(extension)?.format;
}

/** Admit local native Office package data as an embedded object. */
export async function readEmbeddableOfficePackage(
	bytes: ArrayBuffer | Uint8Array,
	extension: string,
): Promise<EmbeddableOfficePackage> {
	const candidate = findOfficePackageFormat(extension);
	if (!candidate) throw new Error('Unsupported object type');
	const zip = await openZip(bytes, {
		label: 'Embedded Office object',
		maxInputBytes: MAX_LOCAL_OBJECT_BYTES,
	});
	if (!zip.file('[Content_Types].xml') || !zip.file(candidate.mainPart)) {
		throw new Error('The object does not match its file type');
	}
	return candidate.format;
}
