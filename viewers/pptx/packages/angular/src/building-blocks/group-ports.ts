import { collectUsedFonts } from 'ooxml-ui/pptx';
import type { BuildingBlockGroupPorts } from 'ooxml-ui/pptx';
import type { PptxCustomProperty } from 'pptx-viewer-core';

import type { LoadNoticesService } from '../viewer/load-notices.service';
import type { ViewerCollaborationShellService } from '../viewer/viewer-collaboration-shell.service';
import type { ViewerCustomShowsService } from '../viewer/viewer-custom-shows.service';
import type { ViewerDialogsService } from '../viewer/viewer-dialogs.service';

export interface AngularGroupSource {
	shell: ViewerCollaborationShellService;
	dialogs: ViewerDialogsService;
	customShows: ViewerCustomShowsService;
	notices: LoadNoticesService;
	/** The bytes the blocks loaded, for the signature list. */
	content(): ArrayBuffer | Uint8Array | null;
	setShortcutsOpen(open: boolean): void;
	translate(key: string): string;
}

/** The Angular ports of the shared dialog groups, over the shell's signals. */
export function createAngularGroupPorts(source: AngularGroupSource): BuildingBlockGroupPorts {
	const { shell, dialogs, customShows } = source;
	const { loader, editor } = shell;
	const markDirty = () => editor.dirty.set(true);
	const setCustom = (next: PptxCustomProperty[]) => loader.customProperties.set([...next]);
	return {
		passwordProtection: {
			input: () => ({
				openPasswordSet: dialogs.isPasswordProtected(),
				recommendation: { kind: source.notices.recommendation().kind },
				hasModifyVerifier: (loader.parsedData()?.modifyVerifier ?? undefined) !== undefined,
				customProperties: loader.customProperties(),
			}),
			customProperties: () => loader.customProperties(),
			applyOpenPassword: (password) => {
				dialogs.presentationPassword.set(password);
				dialogs.isPasswordProtected.set(password !== null);
			},
			applyCustomProperties: setCustom,
			markDirty,
		},
		documentProperties: {
			current: () => ({
				core: loader.coreProperties() ?? {},
				custom: loader.customProperties(),
				app: loader.appProperties() ?? {},
			}),
			applyUpdate: (update) => {
				if (update.core) {
					loader.coreProperties.update((current) => ({ ...current, ...update.core }));
				}
				if (update.app) {
					loader.appProperties.update((current) => ({ ...current, ...update.app }));
				}
				if (update.custom) {
					setCustom(update.custom);
				}
				markDirty();
			},
		},
		fontEmbedding: {
			input: () => ({
				embedFonts: loader.embedFonts(),
				embeddedFamilies: loader.embeddedFonts().map((font) => font.name),
				usedFamilies: collectUsedFonts(editor.slides()),
				descriptor: loader.fontEmbedding(),
			}),
			apply: (enabled) => loader.embedFonts.set(enabled),
			markDirty,
		},
		digitalSignatures: {
			signed: () => loader.hasDigitalSignatures(),
			count: () => loader.digitalSignatureCount(),
			content: () => source.content(),
			markDirty,
		},
		setUpSlideShow: {
			presentationProperties: () => loader.presentationProperties(),
			slideCount: () => editor.slides().length,
			customShows: () => loader.customShows(),
			save: (settings) => {
				loader.presentationProperties.set(settings);
				markDirty();
			},
		},
		customShows: {
			customShows: () => loader.customShows(),
			slides: () => editor.slides(),
			activeId: () => customShows.activeId(),
			commit: (shows, slides) => {
				customShows.onSaveDialog({ shows, clearedSlides: slides });
				markDirty();
			},
			// The blocks mount no slide show overlay: the show is selected and its first slide shown.
			start: (id) => {
				customShows.activeId.set(id);
				shell.activeSlideIndex.set(customShows.showFirstIndex());
			},
		},
		shortcuts: {
			open: () => dialogs.showShortcuts(),
			setOpen: (open) => source.setShortcutsOpen(open),
			toggle: () => source.setShortcutsOpen(!dialogs.showShortcuts()),
			translate: (key) => source.translate(key),
		},
	};
}
