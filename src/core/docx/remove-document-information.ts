// Inspect Document's Remove All, per category, matching what Word 16 saves after each:
// - comments: every comment and its anchors go, and every tracked change is accepted
//   (`WdRemoveDocInfoType` 1 and 2);
// - properties: document properties, personal information and the attached template
//   (types 8, 4 and 9): authors become "Author", initials "A", dates are dropped, and
//   `w:removePersonalInformation` / `w:removeDateAndTime` are set so later saves keep them out;
// - customXml: the custom XML data parts;
// - headersFooters: every header and footer story is emptied (the parts stay, as in Word);
// - hiddenText: runs of hidden (`w:vanish`) text are deleted.
import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	HeaderFooterSlots,
	Paragraph,
} from './model';
import { mapDocumentParagraphs } from './document-paragraphs';
import { acceptAllRevisions } from './revision-commands';
import { anonymizeReviewers } from './personal-information';
import { isHiddenRun, type InspectionCategory } from './inspect-document';
import type { PackageInformation } from './package-information';

const DEFAULT_TEMPLATE = 'Normal.dotm';

function withoutComments(model: DocumentModel): DocumentModel {
	const { comments: _comments, ...rest } = mapDocumentParagraphs(model, (paragraph) =>
		paragraph.runs.some((run) => run.commentIds)
			? {
					...paragraph,
					runs: paragraph.runs.map(({ commentIds: _ids, ...run }) => run),
				}
			: paragraph,
	);
	return rest;
}

function withoutProperties(model: DocumentModel): DocumentModel {
	const { properties: _properties, packageInformation, ...rest } = anonymizeReviewers(model, true);
	const info: PackageInformation = {};
	// The custom XML parts are another category; the template falls back to Normal as in Word.
	if (packageInformation?.customXmlParts) info.customXmlParts = packageInformation.customXmlParts;
	if (packageInformation?.bibliography) info.bibliography = packageInformation.bibliography;
	if (packageInformation?.template) info.template = DEFAULT_TEMPLATE;
	info.removePersonalInformation = true;
	info.removeDateAndTime = true;
	const packageProperties = withoutPersonalProperties(rest.packageProperties);
	const { packageProperties: _loaded, ...withoutLoaded } = rest;
	return {
		...withoutLoaded,
		...(packageProperties ? { packageProperties } : {}),
		packageInformation: info,
	};
}

/**
 * What DOCPROPERTY fields read after the properties are removed: Word keeps the creation and
 * modified dates and the revision number, and clears the rest (company, manager, custom...).
 */
function withoutPersonalProperties(
	loaded: DocumentModel['packageProperties'],
): DocumentModel['packageProperties'] {
	if (!loaded) return undefined;
	const { created, modified, revision } = loaded.core;
	const { application, appVersion, totalTime } = loaded.app;
	return {
		core: {
			...(created !== undefined ? { created } : {}),
			...(modified !== undefined ? { modified } : {}),
			...(revision !== undefined ? { revision } : {}),
		},
		app: {
			...(application !== undefined ? { application } : {}),
			...(appVersion !== undefined ? { appVersion } : {}),
			...(totalTime !== undefined ? { totalTime } : {}),
		},
		custom: [],
	};
}

function withoutCustomXml(model: DocumentModel): DocumentModel {
	if (!model.packageInformation?.customXmlParts && !model.packageInformation?.bibliography)
		return model;
	const { customXmlParts: _parts, bibliography: _bibliography, ...info } = model.packageInformation;
	const { packageInformation: _info, ...rest } = model;
	return Object.keys(info).length ? { ...rest, packageInformation: info } : rest;
}

/** One empty paragraph in place of a story's content, keeping its first paragraph's formatting. */
function emptied(blocks: Block[], id: string): Block[] {
	const first = blocks.find((block): block is Paragraph => block.type === 'paragraph');
	if (!first) return [{ type: 'paragraph', id, runs: [{ text: '' }] }];
	const {
		markRevision: _mark,
		formatRevision: _format,
		bookmarks: _bookmarks,
		permissions: _permissions,
		contentControl: _control,
		dropCap: _dropCap,
		...kept
	} = first;
	return [{ ...kept, runs: [{ text: '' }] }];
}

function withoutHeadersFooters(model: DocumentModel): DocumentModel {
	if (!model.sections?.some((section) => section.headers || section.footers)) return model;
	let next = 0;
	const clear = (slots: HeaderFooterSlots | undefined): HeaderFooterSlots | undefined => {
		if (!slots) return slots;
		const out: HeaderFooterSlots = {};
		for (const [slot, content] of Object.entries(slots) as Array<
			[keyof HeaderFooterSlots, HeaderFooterContent | undefined]
		>)
			if (content)
				out[slot] = { ...content, blocks: emptied(content.blocks, `hf-inspect-${++next}`) };
		return out;
	};
	return {
		...model,
		sections: model.sections.map((section) => {
			const headers = clear(section.headers);
			const footers = clear(section.footers);
			return {
				...section,
				...(headers ? { headers } : {}),
				...(footers ? { footers } : {}),
			};
		}),
	};
}

function withoutHiddenText(model: DocumentModel): DocumentModel {
	return mapDocumentParagraphs(model, (paragraph) => {
		if (!paragraph.runs.some(isHiddenRun)) return paragraph;
		const runs = paragraph.runs.filter((run) => !isHiddenRun(run));
		return { ...paragraph, runs: runs.length ? runs : [{ text: '' }] };
	});
}

const REMOVERS: Record<InspectionCategory, (model: DocumentModel) => DocumentModel> = {
	comments: (model) => acceptAllRevisions(withoutComments(model)),
	properties: withoutProperties,
	customXml: withoutCustomXml,
	headersFooters: withoutHeadersFooters,
	hiddenText: withoutHiddenText,
};

/**
 * The model with everything Inspect Document reports for `categories` removed. Saving it writes
 * what Word writes after Remove All; package parts that only held the information are deleted.
 */
export function removeDocumentInformation(
	model: DocumentModel,
	categories: readonly InspectionCategory[],
): DocumentModel {
	let next = model;
	for (const category of categories) next = REMOVERS[category](next);
	return next;
}
