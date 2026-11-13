import { collectUsedFonts, describeFontEmbedding } from 'ooxml-ui/pptx';
import type { BuildingBlockGroupPorts } from 'ooxml-ui/pptx';

import type { PptxViewer } from '../viewer/PptxViewer';
import type { ViewerState } from '../viewer/state';

export interface VanillaGroupSource {
	viewer(): PptxViewer | null;
	/** The bytes the blocks loaded (null for a URL or Blob source), for the signature list. */
	content(): ArrayBuffer | Uint8Array | null;
	shortcutsOpen(): boolean;
	setShortcutsOpen(open: boolean): void;
	translate(key: string): string;
}

/**
 * The vanilla ports of the shared dialog groups (`registerBuildingBlockGroups`), over the
 * viewer's store and editor. Operations that need a mounted canvas do nothing without one; the
 * groups refuse them with `not-ready` first.
 */
export function createVanillaGroupPorts(source: VanillaGroupSource): BuildingBlockGroupPorts {
	const state = (): ViewerState | null => source.viewer()?.store.get() ?? null;
	const set = (patch: Partial<ViewerState>) => source.viewer()?.store.set(patch);
	const markDirty = () => set({ dirty: true });
	const updateProperties = (
		change: Partial<Pick<ViewerState, 'coreProperties' | 'appProperties' | 'customProperties'>>,
	) => {
		const current = state();
		const viewer = source.viewer();
		if (!current || !viewer) {
			return;
		}
		viewer.editor.updateDocumentProperties(
			change.coreProperties ?? current.coreProperties ?? {},
			change.appProperties ?? current.appProperties ?? {},
			change.customProperties ?? current.customProperties,
		);
	};
	return {
		passwordProtection: {
			input: () => ({
				openPasswordSet: state()?.isPasswordProtected ?? false,
				recommendation: { kind: state()?.readOnlyRecommendation?.kind ?? null },
				hasModifyVerifier: state()?.modifyVerifier !== undefined,
				customProperties: state()?.customProperties ?? [],
			}),
			customProperties: () => state()?.customProperties ?? [],
			applyOpenPassword: (password) =>
				set(
					password === null
						? { isPasswordProtected: false, presentationPassword: null }
						: { isPasswordProtected: true, presentationPassword: password },
				),
			applyCustomProperties: (next) => updateProperties({ customProperties: next }),
			markDirty,
		},
		documentProperties: {
			current: () => ({
				core: state()?.coreProperties ?? {},
				custom: state()?.customProperties ?? [],
				app: state()?.appProperties ?? {},
			}),
			applyUpdate: (update) => {
				const current = state();
				updateProperties({
					...(update.core
						? { coreProperties: { ...current?.coreProperties, ...update.core } }
						: {}),
					...(update.app ? { appProperties: { ...current?.appProperties, ...update.app } } : {}),
					...(update.custom ? { customProperties: update.custom } : {}),
				});
			},
		},
		fontEmbedding: {
			input: () => {
				const current = state();
				const families = (current?.embeddedFonts ?? []).map((font) => font.name);
				return {
					embedFonts: current?.embedFonts ?? false,
					embeddedFamilies: families,
					usedFamilies: current ? collectUsedFonts(current.slides) : [],
					descriptor: describeFontEmbedding(families),
				};
			},
			apply: (enabled) => set({ embedFonts: enabled }),
			markDirty,
		},
		digitalSignatures: {
			signed: () => state()?.hasDigitalSignatures ?? false,
			count: () => state()?.digitalSignatureCount ?? 0,
			content: () => source.content(),
			markDirty,
		},
		setUpSlideShow: {
			presentationProperties: () => state()?.presentationProperties ?? {},
			slideCount: () => state()?.slides.length ?? 0,
			customShows: () => state()?.customShows ?? [],
			save: (settings) => source.viewer()?.editor.updatePresentationProperties(settings),
		},
		customShows: {
			customShows: () => state()?.customShows ?? [],
			slides: () => state()?.slides ?? [],
			activeId: () => state()?.activeCustomShowId ?? null,
			commit: (shows, slides) => {
				const editor = source.viewer()?.editor;
				editor?.updateCustomShows(shows);
				if (slides) {
					editor?.commitElementUpdates(slides);
				}
			},
			start: (id) => {
				const viewer = source.viewer();
				if (!viewer) {
					return;
				}
				viewer.store.set({ activeCustomShowId: id });
				viewer.goToSlide(viewer.firstShowSlideIndex());
				void viewer.enterPresentation();
			},
		},
		shortcuts: {
			open: () => source.shortcutsOpen(),
			setOpen: (open) => source.setShortcutsOpen(open),
			toggle: () => source.setShortcutsOpen(!source.shortcutsOpen()),
			translate: (key) => source.translate(key),
		},
	};
}
