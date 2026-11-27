// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { DocumentModel } from './model';
import { WORD_COMPAT_URI } from './compatibility';
import { SETTINGS_CONTENT_TYPE } from './settings';
import { WORD_NS } from './xml';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';

const SETTINGS_PATH = 'word/settings.xml';
const SETTINGS_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';

/** The compatibility mode Word 2013 and later give a new document (ECMA-376 Part 1, 17.15.3). */
export const NEW_DOCUMENT_COMPATIBILITY_MODE = 15;

/**
 * `word/settings.xml` for a document the model created, in CT_Settings order. It holds what a
 * new Word document declares so Word opens it in its own mode rather than compatibility mode:
 * the zoom, the default tab stop when the model sets one, character spacing control, the
 * separator notes when notes exist and `compatibilityMode`. The single-setting writers in
 * `settings.ts` patch it afterwards. Values that depend on the user's locale (`w:themeFontLang`,
 * `w:decimalSymbol`, `w:listSeparator`) and the random `w:rsids` are left out.
 */
export function newSettingsXml(model: DocumentModel): string {
	const compat = (name: string, value: string | number) =>
		`<w:compatSetting w:name="${name}" w:uri="${WORD_COMPAT_URI}" w:val="${value}"/>`;
	const tabStop = model.defaultTabStopTwips;
	return [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
		`<w:settings xmlns:w="${WORD_NS}">`,
		'<w:zoom w:percent="100"/>',
		tabStop && tabStop > 0 ? `<w:defaultTabStop w:val="${Math.round(tabStop)}"/>` : '',
		'<w:characterSpacingControl w:val="doNotCompress"/>',
		model.footnotes?.length
			? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
			: '',
		model.endnotes?.length
			? '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>'
			: '',
		'<w:compat>',
		compat('compatibilityMode', model.compatibilityMode ?? NEW_DOCUMENT_COMPATIBILITY_MODE),
		'</w:compat>',
		'</w:settings>',
	].join('');
}

/**
 * Writes the whole settings part of a new document, with its content type and relationship, when
 * the package has none yet. A package that already has one is left to the single-setting writers.
 */
export async function writeNewSettings(zip: JSZip, model: DocumentModel): Promise<void> {
	if (zip.file(SETTINGS_PATH)) return;
	zip.file(SETTINGS_PATH, newSettingsXml(model));
	await ensureContentTypeOverride(zip, SETTINGS_PATH, SETTINGS_CONTENT_TYPE);
	await ensureDocumentRelationship(zip, SETTINGS_RELATIONSHIP_TYPE, 'settings.xml');
}
