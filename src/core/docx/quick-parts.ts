// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Insert > Quick Parts > Document Property: a plain-text content control bound (`w:dataBinding`) to
// a core or extended property, as Word writes it. The binding names the property store by its
// well-known `storeItemID`; saving writes an edited control text back to that part
// (`data-binding.ts`).
import type { DocumentPropertySet } from '../opc/properties/types';
import type { ContentControl } from './content-control-model';
import { freshControlId } from './content-control-model';
import { fieldProperties } from './doc-property-field';
import type { DocumentModel } from './model';

/** The `storeItemID` Word gives the core properties part (`docProps/core.xml`). */
export const CORE_PROPERTIES_STORE = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
/** The `storeItemID` Word gives the extended properties part (`docProps/app.xml`). */
export const EXTENDED_PROPERTIES_STORE = '{6668398D-A668-4E3E-A5EB-62B293D839F1}';

const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC = 'http://purl.org/dc/elements/1.1/';
const EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';

export type DocumentPropertyId =
	| 'author'
	| 'title'
	| 'subject'
	| 'keywords'
	| 'comments'
	| 'category'
	| 'status'
	| 'company'
	| 'manager';

export interface DocumentPropertyChoice {
	id: DocumentPropertyId;
	/** The English name Word shows in the gallery and on the control's tab. */
	label: string;
	store: 'core' | 'app';
	/** The XML element the property is stored in. */
	element: string;
	/** The namespace of that element. */
	namespace: string;
	value(properties: DocumentPropertySet): string | undefined;
}

/** The properties of Word's Quick Parts > Document Property gallery that this model stores. */
export const DOCUMENT_PROPERTY_CHOICES: readonly DocumentPropertyChoice[] = [
	{
		id: 'author',
		label: 'Author',
		store: 'core',
		element: 'creator',
		namespace: DC,
		value: (p) => p.core.creator,
	},
	{
		id: 'title',
		label: 'Title',
		store: 'core',
		element: 'title',
		namespace: DC,
		value: (p) => p.core.title,
	},
	{
		id: 'subject',
		label: 'Subject',
		store: 'core',
		element: 'subject',
		namespace: DC,
		value: (p) => p.core.subject,
	},
	{
		id: 'keywords',
		label: 'Keywords',
		store: 'core',
		element: 'keywords',
		namespace: CP,
		value: (p) => p.core.keywords,
	},
	{
		id: 'comments',
		label: 'Comments',
		store: 'core',
		element: 'description',
		namespace: DC,
		value: (p) => p.core.description,
	},
	{
		id: 'category',
		label: 'Category',
		store: 'core',
		element: 'category',
		namespace: CP,
		value: (p) => p.core.category,
	},
	{
		id: 'status',
		label: 'Status',
		store: 'core',
		element: 'contentStatus',
		namespace: CP,
		value: (p) => p.core.contentStatus,
	},
	{
		id: 'company',
		label: 'Company',
		store: 'app',
		element: 'Company',
		namespace: EP,
		value: (p) => p.app.company,
	},
	{
		id: 'manager',
		label: 'Manager',
		store: 'app',
		element: 'Manager',
		namespace: EP,
		value: (p) => p.app.manager,
	},
];

const escapeAttribute = (text: string) =>
	text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** `w:dataBinding` of `choice`, the way Word writes it (prefixes `ns0` and `ns1`). */
export function propertyBinding(choice: DocumentPropertyChoice): {
	prefixMappings: string;
	xpath: string;
	storeItemId: string;
} {
	if (choice.store === 'app')
		return {
			prefixMappings: `xmlns:ns0='${EP}'`,
			xpath: `/ns0:Properties[1]/ns0:${choice.element}[1]`,
			storeItemId: EXTENDED_PROPERTIES_STORE,
		};
	const prefix = choice.namespace === DC ? 'ns0' : 'ns1';
	return {
		prefixMappings: `xmlns:ns0='${DC}' xmlns:ns1='${CP}'`,
		xpath: `/ns1:coreProperties[1]/${prefix}:${choice.element}[1]`,
		storeItemId: CORE_PROPERTIES_STORE,
	};
}

export function documentPropertyChoice(id: string): DocumentPropertyChoice | undefined {
	return DOCUMENT_PROPERTY_CHOICES.find((choice) => choice.id === id);
}

/** The property's current value in `model`, with the core fields edited in the document applied. */
export function documentPropertyValue(model: DocumentModel, id: DocumentPropertyId): string {
	const choice = documentPropertyChoice(id);
	return choice?.value(fieldProperties(model)) ?? '';
}

/**
 * The content control Word inserts for Document Property `id`: a text control titled with the
 * property's name, bound to its property part, showing its value (or a bracketed placeholder while
 * the property is empty).
 */
export function documentPropertyControl(
	id: DocumentPropertyId,
	used: ReadonlySet<string>,
	label?: string,
): { control: ContentControl; value: string } | undefined {
	const choice = documentPropertyChoice(id);
	if (!choice) return undefined;
	const controlId = freshControlId(used);
	const binding = propertyBinding(choice);
	const name = label ?? choice.label;
	const title = escapeAttribute(name);
	const sourceXml =
		`<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:alias w:val="${title}"/><w:tag w:val="${title}"/><w:id w:val="${controlId}"/>` +
		`<w:dataBinding w:prefixMappings="${escapeAttribute(binding.prefixMappings)}" w:xpath="${escapeAttribute(binding.xpath)}" w:storeItemID="${binding.storeItemId}"/>` +
		`<w:text/></w:sdtPr>`;
	return {
		control: {
			id: controlId,
			level: 'inline',
			kind: 'text',
			alias: name,
			tag: name,
			dataBound: true,
			sourceXml,
		},
		value: '',
	};
}
