import { OLE_OBJECT_NAME_SECTIONS, resolveOleSummaryRows } from 'ooxml-ui/pptx';
import type { OlePptxElement, PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { createInspectorFields } from './inspector-fields';
import { openOleEditorDialog } from './ole-editor-dialog';
import type { InspectorHandlers, InspectorState } from './types';

export interface OlePropertiesSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * OLE object summary (type, file name, link status) plus the Object Name
 * editor. A browser cannot run the native application that owns an embedded
 * OLE object, so the object itself stays read-only. Its Object Name IS
 * editable: `p:oleObj/@name` (ECMA-376 SS13.3.4) already parses, saves, and
 * syncs via collaboration, and shared's `getOleDisplayName` / `getOleAriaLabel`
 * already read it, so this field was the only piece missing to make it a
 * real, round-tripping edit.
 */
export function createOlePropertiesSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): OlePropertiesSection {
	const el = section(t('pptx.ole.title'));
	el.classList.add('pptxv-ole-info');

	// A browser cannot run the native application that owns an embedded OLE
	// object, but its recovered payload (spreadsheet cells, document
	// paragraphs, nested-deck slide titles) IS editable in place; see
	// `ole-editor-dialog.ts`. Gated the same as every other per-element edit
	// action in this inspector: not locked, and (OLE-specific) not a link,
	// since there is no local payload to edit for a linked object.
	let currentOle: OlePptxElement | undefined;
	const editButton = doc.createElement('button');
	editButton.type = 'button';
	editButton.textContent = t('pptx.ole.editContent');
	editButton.addEventListener('click', () => {
		if (currentOle) {
			openOleEditorDialog(doc, t, currentOle, {
				onUpdateElement: (patch) => handlers.setOleContent(patch),
				regions: handlers.getDialogRegions?.(),
			});
		}
	});
	el.appendChild(editButton);

	// The Object Name: the shared descriptor, drawn by the generic renderer.
	const name = createInspectorFields(doc, t, {
		sections: OLE_OBJECT_NAME_SECTIONS,
		onPatch: (patch) => handlers.patchSelectedElement(patch),
	});
	el.appendChild(name.el);
	// The read-only rows (type, file name, linked or embedded): `resolveOleSummaryRows`.
	const rows = createEl(doc, 'div', 'pptxv-ole-rows');
	el.appendChild(rows);

	return {
		el,
		update(state) {
			el.hidden = !state.isOle;
			if (!state.isOle) {
				return;
			}
			currentOle = state.oleElement;
			editButton.hidden = state.isLocked || state.oleIsLinked;
			if (state.oleElement) {
				name.update(state.oleElement, false, state.oleElement.id);
			}
			rows.replaceChildren(
				...resolveOleSummaryRows(state.oleElement ?? ({ type: 'shape' } as PptxElement)).map(
					(summary) => {
						const row = createEl(doc, 'div', 'pptxv-ole-row');
						const label = createEl(doc, 'span');
						label.textContent = t(summary.labelKey);
						const value = createEl(doc, 'span', 'pptxv-ole-value');
						value.textContent = summary.valueKey ? t(summary.valueKey) : (summary.value ?? '');
						value.title = summary.value ?? '';
						if (summary.id === 'linkStatus') {
							value.classList.add('pptxv-ole-badge');
							value.classList.toggle('is-linked', state.oleIsLinked);
						}
						row.append(label, value);
						return row;
					},
				),
			);
		},
	};
}
