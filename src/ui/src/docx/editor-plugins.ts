import type { DocumentModel } from 'ooxml-core/docx';
import { history } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import { editorKeymap } from './editor-commands';
import {
	contentControlPlugin,
	editRestrictionPlugin,
	fieldGuardPlugin,
	forcesTrackChanges,
	formFieldGuardPlugin,
	positionalTabInputPlugin,
	symbolFontPlugin,
	tableStructureGuardPlugin,
} from 'ooxml-core/docx/ui';
import { contentControlViewPlugin } from './content-control-view';
import { formFieldViewPlugin } from './form-field-view';
import { editableRegionViewPlugin } from './editable-region-view';
import { tabStopsPlugin } from './tab-stops-view';
import { noteNumberingPlugin } from './note-commands';
import { paragraphStylesPlugin } from './paragraph-styles';
import { reviewDisplayPlugin, type ReviewDisplayMode } from './review-display';
import { runStylesPlugin } from './run-styles';
import { sectionBreaksPlugin } from './section-commands';
import { trackChangesPlugin } from './track-changes-mode';
import { reviewRunMarksPlugin } from './review-run-marks';
import { wordPastePlugin } from './word-paste';
import { autoCorrectPlugin } from './autocorrect-plugin';
import { PasteController, pasteOptionsPlugin } from './paste-options';
import { pasteOptionsButtonPlugin, type PasteOptionsButtonHost } from './paste-options-button';

/** Word's paste: its HTML read as Word means it, Paste Options after it, Paste Special. */
function pastePlugins(host: BodyPluginHost): Plugin[] {
	const controller = new PasteController();
	return [
		wordPastePlugin(
			() => host.model(),
			() => (controller.option === 'merge' ? 'merge' : 'source'),
		),
		pasteOptionsPlugin(controller),
		...(host.pasteOptions ? [pasteOptionsButtonPlugin(host.pasteOptions)] : []),
		keymap({ 'Mod-Alt-v': () => (host.pasteSpecial?.(), Boolean(host.pasteSpecial)) }),
	];
}

/** What the body editor's plugins read from the editor element. */
export interface BodyPluginHost {
	model(): DocumentModel;
	/** Object URL of a picture bullet's image part, when its bytes are loaded. */
	pictureUrl?(partName: string, contentType: string): string | undefined;
	reviewAuthor(): string;
	highlightEditableRegions?(): boolean;
	reviewDisplayMode(): ReviewDisplayMode;
	insertNote(kind: 'footnote' | 'endnote'): void;
	showSearch(): void;
	/** Paste Special (Ctrl+Alt+V). */
	pasteSpecial?(): void;
	/** Where the Paste Options button is drawn; without it pastes are still tracked. */
	pasteOptions?: PasteOptionsButtonHost;
	/** Controller and collaboration plugins, in the order they apply. */
	extraPlugins: Plugin[];
	collaborationPlugins: Plugin[];
	yjs?: boolean;
	revisionIds?: ((kind: string) => string) | undefined;
}

/** The main body editor's plugins: history, structure guards, style decorations, review and keys. */
export function bodyPlugins(host: BodyPluginHost): Plugin[] {
	return [
		...(host.yjs ? [] : [history()]),
		...host.extraPlugins,
		...pastePlugins(host),
		noteNumberingPlugin(),
		sectionBreaksPlugin(),
		fieldGuardPlugin(),
		tableStructureGuardPlugin(),
		contentControlPlugin(),
		contentControlViewPlugin(),
		formFieldGuardPlugin(() => host.model().protection),
		editRestrictionPlugin(
			() => host.model().protection,
			() => host.reviewAuthor(),
		),
		editableRegionViewPlugin(
			() => host.model().protection,
			() => host.reviewAuthor(),
			() => host.highlightEditableRegions?.() ?? true,
		),
		formFieldViewPlugin(() => host.model()),
		symbolFontPlugin(),
		positionalTabInputPlugin(),
		autoCorrectPlugin(() => host.model()),
		reviewRunMarksPlugin(() => host.reviewDisplayMode()),
		tabStopsPlugin(),
		keymap({
			'Mod-Alt-f': () => (host.insertNote('footnote'), true),
			'Mod-Alt-d': () => (host.insertNote('endnote'), true),
		}),
		runStylesPlugin(
			() => host.model(),
			() => host.reviewDisplayMode(),
		),
		paragraphStylesPlugin(
			() => host.model(),
			() => host.reviewDisplayMode(),
			(partName, contentType) => host.pictureUrl?.(partName, contentType),
		),
		trackChangesPlugin(
			() => host.reviewAuthor(),
			() => Boolean(host.model().trackChanges) || forcesTrackChanges(host.model().protection),
			host.revisionIds,
		),
		reviewDisplayPlugin(() => host.reviewDisplayMode()),
		editorKeymap(
			() => host.showSearch(),
			() => host.model(),
		),
		...host.collaborationPlugins,
	];
}
