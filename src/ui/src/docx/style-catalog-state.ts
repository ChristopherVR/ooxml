import type { DocumentModel, StyleCatalogs } from 'ooxml-core/docx';
import type { Node } from 'prosemirror-model';

/** Catalog changes are root attributes so undo and redo restore the whole edit atomically. */
export function withStyleCatalogs(model: DocumentModel, doc: Node): DocumentModel {
	const catalogs: StyleCatalogs | null = doc.attrs.styleCatalogs;
	if (!catalogs) return model;
	if (
		catalogs.paragraphStyles === model.paragraphStyles &&
		catalogs.characterStyles === model.characterStyles &&
		catalogs.tableStyles === model.tableStyles
	)
		return model;
	const { paragraphStyles: _p, characterStyles: _r, tableStyles: _t, ...rest } = model;

	// List commands maintain heading-to-list associations independently of style authoring.
	const paragraphStyles = catalogs.paragraphStyles;
	const current = model.paragraphStyles;
	if (!paragraphStyles || !current) return { ...rest, ...catalogs };
	let styles = paragraphStyles.styles;
	for (const [id, style] of Object.entries(paragraphStyles.styles)) {
		const existing = current.styles[id];
		if (!existing || existing.numbering === style.numbering) continue;
		if (styles === paragraphStyles.styles) styles = { ...styles };
		const { numbering: _numbering, ...rest } = style;
		styles[id] = { ...rest, ...(existing.numbering && { numbering: existing.numbering }) };
	}
	return {
		...rest,
		...catalogs,
		paragraphStyles:
			styles === paragraphStyles.styles ? paragraphStyles : { ...paragraphStyles, styles },
	};
}
