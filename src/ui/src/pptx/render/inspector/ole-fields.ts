/**
 * The inspector's OLE object card, written once: the Object Name
 * (`p:oleObj/@name`) as a field descriptor every binding draws with its one
 * generic renderer, and {@link resolveOleSummaryRows}, the read-only rows
 * (object type, file name, linked or embedded) a binding maps onto its
 * template. The Edit Content action and its dialog stay native.
 *
 * The name is written as it is typed (a blank box removes it). Trimming each
 * keystroke, as the bindings did, swallowed the space a user had just typed
 * between two words.
 *
 * @module render/inspector/ole-fields
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getOleObjectTypeLabel } from 'ooxml-core/pptx';

import type { InspectorSectionDescriptor } from './inspector-field-model';

export type OleSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;

/** The Object Name field (untitled: the OLE card heads it). */
export const OLE_OBJECT_NAME_SECTION: OleSection = {
	id: 'oleObject',
	fields: [
		{
			id: 'oleName',
			kind: 'text',
			labelKey: 'pptx.ole.objectName',
			placeholderKey: 'pptx.ole.objectNamePlaceholder',
			wide: true,
			testId: 'ole-object-name',
			read: (element) => (element.type === 'ole' ? (element.oleName ?? '') : ''),
			write: (element, value) =>
				element.type === 'ole'
					? ({ oleName: String(value).trim() ? String(value) : undefined } as Partial<PptxElement>)
					: {},
		},
	],
};

/** The Object Name field as a section list, what a generic renderer takes. */
export const OLE_OBJECT_NAME_SECTIONS: readonly OleSection[] = [OLE_OBJECT_NAME_SECTION];

/** One read-only row of the OLE card. */
export interface OleSummaryRow {
	id: 'type' | 'fileName' | 'linkStatus';
	/** Translation key of the row's caption. */
	labelKey: string;
	/** The value as text, when it is not a dictionary entry. */
	value?: string;
	/** Translation key of the value, when it is one. */
	valueKey?: string;
}

/** The read-only rows of `element`'s OLE card (none for another kind). */
export function resolveOleSummaryRows(element: PptxElement): OleSummaryRow[] {
	if (element.type !== 'ole') {
		return [];
	}
	const rows: OleSummaryRow[] = [
		{ id: 'type', labelKey: 'pptx.ole.type', value: getOleObjectTypeLabel(element.oleObjectType) },
	];
	if (element.fileName) {
		rows.push({ id: 'fileName', labelKey: 'pptx.ole.fileName', value: element.fileName });
	}
	rows.push({
		id: 'linkStatus',
		labelKey: 'pptx.ole.linkStatus',
		valueKey: element.isLinked ? 'pptx.ole.linked' : 'pptx.ole.embedded',
	});
	return rows;
}
