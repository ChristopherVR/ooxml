import type { DocumentModel } from 'ooxml-core/docx';

/**
 * The document properties and package information, kept as a JSON document attribute
 * (`documentInformation`) so that removing them (File > Info > Inspect Document) is one undoable
 * step with the rest of the document.
 */
type DocumentInformation = Pick<DocumentModel, 'properties' | 'packageInformation'>;

export function documentInformationJson(model: DocumentModel): string {
	const info: DocumentInformation = {};
	if (model.properties) info.properties = model.properties;
	if (model.packageInformation) info.packageInformation = model.packageInformation;
	return JSON.stringify(info);
}

/**
 * The fields the attribute holds, or undefined to keep the prior model's (a document built
 * without the attribute, such as one received from a collaboration room).
 */
export function readDocumentInformation(value: unknown): DocumentInformation | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		const parsed = JSON.parse(value) as DocumentInformation;
		return parsed && typeof parsed === 'object' ? parsed : undefined;
	} catch {
		return undefined;
	}
}
