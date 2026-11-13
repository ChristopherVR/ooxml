import { collectUsedFonts, resolveAuthoredCustomShowId } from 'ooxml-ui/pptx';
import type { BuildingBlockGroupPorts } from 'ooxml-ui/pptx';

import type { ViewerStateBag } from '../viewer/state/create-viewer-state-types';

export interface SvelteGroupSource {
	state: ViewerStateBag;
	/** The bytes the blocks loaded, for the signature list. */
	content(): ArrayBuffer | Uint8Array | null | undefined;
	/** Whether the deck's open password is set now (the file's, until the host clears it). */
	openPasswordSet(): boolean;
	setOpenPasswordSet(set: boolean): void;
}

/**
 * The Svelte ports of the shared dialog groups (`registerBuildingBlockGroups`), over the viewer
 * state's editor and options. Operations that need an editable deck are refused by the groups
 * first; the editor itself ignores them while not editable.
 */
export function createSvelteGroupPorts(source: SvelteGroupSource): BuildingBlockGroupPorts {
	const { state } = source;
	const { editor, loader, parityUi } = state;
	const markDirty = () => {
		editor.dirty = true;
	};
	const updateProperties = (change: {
		core?: typeof editor.coreProperties;
		app?: typeof editor.appProperties;
		custom?: typeof editor.customProperties;
	}) =>
		editor.updateDocumentProperties(
			change.core ?? editor.coreProperties ?? {},
			change.app ?? editor.appProperties ?? {},
			change.custom ?? editor.customProperties,
			change.core ?? null,
		);
	return {
		passwordProtection: {
			input: () => ({
				openPasswordSet: source.openPasswordSet(),
				recommendation: { kind: state.readOnlyRec.recommendation.kind },
				hasModifyVerifier: loader.modifyVerifier !== undefined,
				customProperties: editor.customProperties,
			}),
			customProperties: () => editor.customProperties,
			applyOpenPassword: (password) => {
				if (password === null) {
					editor.clearSavePassword();
				} else {
					editor.setSavePassword(password);
				}
				source.setOpenPasswordSet(password !== null);
			},
			applyCustomProperties: (next) => updateProperties({ custom: next }),
			markDirty,
		},
		documentProperties: {
			current: () => ({
				core: editor.coreProperties ?? {},
				custom: editor.customProperties,
				app: editor.appProperties ?? {},
			}),
			applyUpdate: (update) =>
				updateProperties({
					...(update.core ? { core: { ...editor.coreProperties, ...update.core } } : {}),
					...(update.app ? { app: { ...editor.appProperties, ...update.app } } : {}),
					...(update.custom ? { custom: update.custom } : {}),
				}),
		},
		fontEmbedding: {
			input: () => ({
				embedFonts: editor.embedFonts,
				embeddedFamilies: [...editor.embeddedFontFamilies],
				usedFamilies: collectUsedFonts(editor.slides),
				descriptor: editor.fontEmbedding,
			}),
			apply: (enabled) => {
				editor.embedFonts = enabled;
			},
			markDirty,
		},
		digitalSignatures: {
			signed: () => loader.hasDigitalSignatures,
			count: () => loader.digitalSignatureCount,
			content: () => source.content(),
			markDirty,
		},
		setUpSlideShow: {
			presentationProperties: () => editor.presentationProperties,
			slideCount: () => editor.slides.length,
			customShows: () => editor.customShows,
			save: (settings) => {
				editor.presentationMetadata.updatePresentationProperties(settings);
				parityUi.activeCustomShowId =
					resolveAuthoredCustomShowId(settings, editor.customShows) ?? null;
			},
		},
		customShows: {
			customShows: () => editor.customShows,
			slides: () => editor.slides,
			activeId: () => parityUi.activeCustomShowId,
			commit: (shows, slides) => {
				editor.presentationMetadata.updateCustomShows(shows);
				if (slides) {
					editor.commitSlides([...slides]);
				}
			},
			start: (id) => {
				parityUi.activeCustomShowId = id;
				state.presentation.firstSlide();
				state.onFullscreenToggle();
			},
		},
		shortcuts: {
			open: () => parityUi.shortcutsOpen,
			setOpen: (open) => {
				parityUi.shortcutsOpen = open;
			},
			toggle: () => {
				parityUi.shortcutsOpen = !parityUi.shortcutsOpen;
			},
			translate: (key) => state.t(key),
		},
	};
}
