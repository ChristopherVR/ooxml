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
	const { bibliography: _sources, ...local } = model.packageInformation ?? {};
	if (Object.keys(local).length) info.packageInformation = local;
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

/** The room shares the list XML. Each peer keeps the part path in its own source package. */
export function withBibliographySources<T extends Partial<DocumentModel>>(
	model: T,
	xml: unknown,
	localPart: string | undefined,
): T {
	const { bibliography, ...local } = model.packageInformation ?? {};
	if (typeof xml === 'string') {
		if (xml === bibliography?.xml) return model;
		return {
			...model,
			packageInformation: {
				...local,
				bibliography: { ...(localPart ? { part: localPart } : {}), xml },
			},
		};
	}
	if (xml !== null || !bibliography) return model;
	const { packageInformation: _info, ...rest } = model;
	return Object.keys(local).length ? { ...model, packageInformation: local } : (rest as T);
}
