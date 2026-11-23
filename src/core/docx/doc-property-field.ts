// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// DOCPROPERTY: a document property by name (ECMA-376 Part 1, 17.16.5.17), read from the core,
// extended and custom property parts through the shared `opc/properties` model.
import type {
	AppProperties,
	CoreProperties,
	CustomProperty,
	DocumentPropertySet,
} from '../opc/properties/types';
import { applyFieldFormat, fieldDateResult, parseFieldInstruction } from './field-format';
import type { DocumentModel } from './model';

/** Word's results when a DOCPROPERTY names nothing it knows (English Word's wording). */
export const UNKNOWN_PROPERTY = 'Error! Unknown document property name.';
export const NO_PROPERTY_NAME = 'Error! No property name supplied.';

/** The text value, number or date a property gives, or `stale` when it cannot be computed. */
type PropertyValue =
	| { kind: 'text'; text: string }
	| { kind: 'date'; date: Date }
	| { kind: 'stale' }
	| undefined;

const text = (value: string | number | undefined): PropertyValue => ({
	kind: 'text',
	text: value === undefined ? '' : String(value),
});

/** A W3CDTF date; a property that holds no readable date cannot be drawn. */
function date(value: string | undefined): PropertyValue {
	if (value === undefined) return text(undefined);
	const parsed = new Date(value);
	return Number.isNaN(parsed.getTime()) ? { kind: 'stale' } : { kind: 'date', date: parsed };
}

/**
 * Word's built-in property names (any case). The document statistics (`Pages`, `Words`...) are
 * what Word counts as it lays the document out; the saved counts are stale, so those fields keep
 * their results, as does `TotalEditingTime`, which counts the time the document is open.
 */
function builtIn(name: string, core: CoreProperties, app: AppProperties): PropertyValue {
	switch (name) {
		case 'title':
			return text(core.title);
		case 'subject':
			return text(core.subject);
		case 'author':
			return text(core.creator);
		case 'keywords':
			return text(core.keywords);
		case 'comments':
			return text(core.description);
		case 'lastsavedby':
			return text(core.lastModifiedBy);
		case 'revisionnumber':
			return text(core.revision);
		case 'category':
			return text(core.category);
		case 'manager':
			return text(app.manager);
		case 'company':
			return text(app.company);
		case 'hyperlinkbase':
			return text(app.hyperlinkBase);
		case 'template':
			return text(app.template);
		case 'nameofapplication':
			return text(app.application);
		case 'security':
			return text(app.docSecurity ?? 0);
		case 'createtime':
			return date(core.created);
		case 'lastsavedtime':
			return date(core.modified);
		case 'lastprinted':
			return date(core.lastPrinted);
		case 'pages':
		case 'words':
		case 'characters':
		case 'characterswithspaces':
		case 'lines':
		case 'paragraphs':
		case 'bytes':
		case 'totaleditingtime':
			return { kind: 'stale' };
		default:
			return undefined;
	}
}

/** A custom property's value: Word shows a yes/no property as `Y` or `N`. */
function custom(property: CustomProperty): PropertyValue {
	switch (property.type) {
		case 'lpwstr':
			return text(property.value);
		case 'i4':
		case 'r8':
			return text(property.value);
		case 'bool':
			return text(property.value ? 'Y' : 'N');
		case 'filetime':
			return date(property.value);
		default:
			return { kind: 'stale' };
	}
}

/** Word document-information fields that mirror a built-in DOCPROPERTY (ECMA-376 17.16.5). */
export const DOC_INFO_FIELD_MAP: Readonly<Record<string, string>> = {
	AUTHOR: 'author',
	TITLE: 'title',
	SUBJECT: 'subject',
	KEYWORDS: 'keywords',
	COMMENTS: 'comments',
	LASTSAVEDBY: 'lastsavedby',
	REVNUM: 'revisionnumber',
	CREATEDATE: 'createtime',
	SAVEDATE: 'lastsavedtime',
	PRINTDATE: 'lastprinted',
};

/** Fields whose optional first argument supplies a literal override text in Word. */
const OVERRIDABLE_INFO_FIELDS = new Set(['AUTHOR', 'TITLE', 'SUBJECT', 'KEYWORDS', 'COMMENTS']);

/**
 * Evaluates a document-information field (`AUTHOR`, `TITLE`, `SUBJECT`, `KEYWORDS`, `COMMENTS`,
 * `LASTSAVEDBY`, `REVNUM`, `CREATEDATE`, `SAVEDATE`, `PRINTDATE`).
 */
export function docInfoFieldResult(
	instr: string,
	properties: DocumentPropertySet,
): string | undefined {
	const field = parseFieldInstruction(instr);
	if (!field) return undefined;
	const propKey = DOC_INFO_FIELD_MAP[field.name];
	if (!propKey) return undefined;
	if (field.args.length > 0 && OVERRIDABLE_INFO_FIELDS.has(field.name)) {
		return applyFieldFormat(field.args[0]!, field.format);
	}
	const raw = rawDocInfoProperty(propKey, properties.core);
	if (raw === undefined) return undefined;
	const value = builtIn(propKey, properties.core, properties.app);
	if (!value || value.kind === 'stale') return undefined;
	if (value.kind === 'date') return fieldDateResult(value.date, field.format);
	return applyFieldFormat(value.text, field.format);
}

function rawDocInfoProperty(key: string, core: CoreProperties): string | undefined {
	switch (key) {
		case 'author':
			return core.creator;
		case 'title':
			return core.title;
		case 'subject':
			return core.subject;
		case 'keywords':
			return core.keywords;
		case 'comments':
			return core.description;
		case 'lastsavedby':
			return core.lastModifiedBy;
		case 'revisionnumber':
			return core.revision;
		case 'createtime':
			return core.created;
		case 'lastsavedtime':
			return core.modified;
		case 'lastprinted':
			return core.lastPrinted;
		default:
			return undefined;
	}
}

/**
 * A DOCPROPERTY result, with its `\*` formats and, for a date, its `\@` picture. Undefined when
 * the field keeps its saved result: another field, an unevaluated switch, a statistic or a date
 * without a picture (Word draws those in the system's regional format).
 */
export function docPropertyResult(
	instr: string,
	properties: DocumentPropertySet,
): string | undefined {
	const field = parseFieldInstruction(instr);
	if (field?.name !== 'DOCPROPERTY') return undefined;
	const name = field.args[0];
	if (name === undefined) return NO_PROPERTY_NAME;
	const key = name.toLowerCase();
	const value =
		builtIn(key, properties.core, properties.app) ??
		(() => {
			const found = properties.custom.find((item) => item.name.toLowerCase() === key);
			return found && custom(found);
		})();
	if (!value) return UNKNOWN_PROPERTY;
	if (value.kind === 'stale') return undefined;
	if (value.kind === 'date') return fieldDateResult(value.date, field.format);
	return applyFieldFormat(value.text, field.format);
}

/** The core fields a document's editable `properties` hold; they win over the parsed part. */
const EDITABLE = [
	'title',
	'subject',
	'creator',
	'keywords',
	'description',
	'lastModifiedBy',
] as const;

/**
 * The properties a DOCPROPERTY reads: those the package was loaded with, with the core fields
 * edited in the document (`model.properties`) taking their place.
 */
export function fieldProperties(model: DocumentModel): DocumentPropertySet {
	const loaded = model.packageProperties;
	const core: CoreProperties = { ...loaded?.core };
	if (model.properties || !loaded)
		for (const key of EDITABLE) {
			const value = model.properties?.[key];
			if (value === undefined) delete core[key];
			else core[key] = value;
		}
	return { core, app: loaded?.app ?? {}, custom: loaded?.custom ?? [] };
}
