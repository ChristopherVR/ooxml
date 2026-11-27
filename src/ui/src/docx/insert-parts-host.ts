import { fieldProperties, glossaryEntryKey, type GlossaryEntry } from 'ooxml-core/docx';
import { closeHistory } from 'prosemirror-history';
import type { EditorCore } from './editor-core';
import type { InsertPartsHost } from './insert-parts-dialogs';
import { focusView } from './focus-view';
import { newPicturePartName } from './picture-commands';
import { insertGlossaryEntry, selectionBlocks } from './quick-parts-commands';
import { insertSignatureLine, insertWebVideo } from './insert-parts-insert';
import { schema } from './schema';
import { emit } from './events';
import { translate, type LocalizationKey } from './localization';

/**
 * What the Quick Parts, Signature Line and Online Video dialogs use of the editor. Building blocks
 * live on the model (`glossary`) and are written to the glossary part on save; a document being
 * edited by several people does not take glossary changes, since that part is not shared.
 */
export function createInsertPartsHost(core: EditorCore): InsertPartsHost {
	const stage = (bytes: Uint8Array, contentType: string): string => {
		const media = core.collab.yjs?.media;
		const partName = media?.partName(contentType) ?? newPicturePartName(contentType);
		const part = { bytes, contentType };
		media?.publish(partName, part);
		core.inserts.pendingMedia.set(partName, part);
		return partName;
	};
	const view = () => core.targetView();
	const editable = () =>
		Boolean(view()?.editable) && core.canEditBody() && !core.loading && !core.readOnly;
	const setGlossary = (entries: GlossaryEntry[]) => {
		const { glossary: _previous, ...rest } = core.model;
		core.model = entries.length ? { ...rest, glossary: entries } : rest;
		core.notifyChange();
		core.refreshControls();
	};
	const reportStoreError = () =>
		emit(
			core.element,
			'document-warning',
			translate(core.locale, 'The shared building blocks could not be saved.' as LocalizationKey),
		);
	return {
		editable,
		focus: () => focusView(view()),
		glossary: () => core.model.glossary ?? [],
		saveGlossaryEntry(entry) {
			if (core.collab.active) return;
			const key = glossaryEntryKey(entry);
			setGlossary([
				...(core.model.glossary ?? []).filter((item) => glossaryEntryKey(item) !== key),
				entry,
			]);
		},
		removeGlossaryEntry(key) {
			if (core.collab.active) return;
			setGlossary((core.model.glossary ?? []).filter((item) => glossaryEntryKey(item) !== key));
		},
		insertEntry(entry) {
			const target = view();
			if (target) insertGlossaryEntry(target, core.model, entry);
			focusView(target);
		},
		selectionBlocks() {
			const target = view();
			return core.collab.active || !target ? undefined : selectionBlocks(target, core.model);
		},
		insertSignatureLine(spec) {
			const target = view();
			if (target)
				void insertSignatureLine(
					target,
					spec,
					stage,
					translate(core.locale, 'Signature line' as LocalizationKey),
				);
		},
		glossaryLocked: () => core.collab.active,
		sharedGlossaryAvailable: () => Boolean(core.buildingBlocks.store),
		sharedGlossary: () => core.buildingBlocks.entries(),
		loadSharedGlossary: () => core.buildingBlocks.load(true),
		saveSharedEntry(entry) {
			core.buildingBlocks.save(entry).catch(() => reportStoreError());
		},
		removeSharedEntry(key) {
			core.buildingBlocks.remove(key).catch(() => reportStoreError());
		},
		report: (message) => emit(core.element, 'document-warning', message),
		insertWebVideo(spec) {
			const target = view();
			if (target) insertWebVideo(target, spec, stage, translate(core.locale, 'Online Video'));
		},
		onlineVideoHosts: () => core.onlineVideoHosts,
		onlineVideoSearch: () => core.onlineVideoSearch,
		properties: () => fieldProperties(core.model),
		insertField(instr, result) {
			const target = view();
			if (!target?.editable) return;
			const mark = schema.marks.field!.create({ instr, simple: true });
			target.dispatch(
				closeHistory(target.state.tr.replaceSelectionWith(schema.text(result, [mark]), false)),
			);
			focusView(target);
		},
	};
}
