import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	hyperlinkEditPatch,
	resetSlideLayoutPath,
	scopeLayoutOptionsToSlide,
	selectionEquation,
} from 'ooxml-ui/pptx';
import type {
	BuildingBlockAvailabilityInput,
	BuildingBlockFindReplacePort,
	BuildingBlockGroupPorts,
	BuildingBlockSessionPort,
} from 'ooxml-ui/pptx';
import { createEditorId } from 'pptx-viewer-core';
import type { PptxElement, PptxLayoutOption } from 'pptx-viewer-core';

import type { BuildingBlocksEngineContext as Context } from './engine-context';
import { mediaElementFromFile } from 'ooxml-ui/pptx';

export type EngineSessionPort = Omit<
	BuildingBlockSessionPort,
	'fileLimits' | 'onRequest' | 'onCancel'
>;

/** The find port over the viewer's own find state (`useFindReplace`), as React's. */
function findPort(c: Context): BuildingBlockFindReplacePort {
	const { find } = c;
	return {
		open: c.findOpen.value,
		setOpen: (open) => {
			c.findOpen.value = open;
		},
		query: find.query.value,
		replacement: find.replacement.value,
		options: {
			matchCase: find.matchCase.value,
			includeNotes: find.includeNotes.value,
			includeChartTitles: find.includeChartTitles.value,
			includeChartData: find.includeChartData.value,
		},
		matches: find.matches.value,
		currentIndex: find.currentMatch.value,
		setQuery: (query) => {
			find.query.value = query;
		},
		setReplacement: (replacement) => {
			find.replacement.value = replacement;
		},
		setOptions: (change) => {
			if (change.matchCase !== undefined) {
				find.matchCase.value = change.matchCase;
			}
			if (change.includeNotes !== undefined) {
				find.includeNotes.value = change.includeNotes;
			}
			if (change.includeChartTitles !== undefined) {
				find.includeChartTitles.value = change.includeChartTitles;
			}
			if (change.includeChartData !== undefined) {
				find.includeChartData.value = change.includeChartData;
			}
		},
		// The query setter searches; assigning it again re-runs the search.
		find: () => {
			const { value } = find.query;
			find.query.value = value;
		},
		navigate: (delta) => (delta > 0 ? find.next() : find.prev()),
		replaceCurrent: find.replaceCurrent,
		replaceAll: find.replaceAll,
	};
}

/** The layouts the active slide may switch to (its master's). */
function scopedLayouts(c: Context): PptxLayoutOption[] {
	return scopeLayoutOptionsToSlide(
		[...c.layoutOptions.value] as PptxLayoutOption[],
		c.activeSlide.value?.layoutPath,
	);
}

/** The availability input of the Vue viewer, read from its refs. */
function availabilityInput(c: Context): BuildingBlockAvailabilityInput {
	const element = c.selectedElements.value[0] ?? null;
	const slide = c.activeSlide.value;
	return {
		ready: !c.loading.value && !c.error.value,
		hasActiveSlide: slide !== undefined,
		canEdit: c.canEditBase(),
		editing: c.canEdit.value && !c.presenting.value && !c.showMasterView.value,
		canInsert: c.canInsert(),
		hasSelection: element !== null,
		selectionHasHyperlink: Boolean(element?.actionClick),
		selectionHasEquation: selectionEquation(element) !== null,
		layoutCount: scopedLayouts(c).length,
		hasCurrentLayout: resetSlideLayoutPath(slide) !== undefined,
		hasHandler: c.handler.value !== null,
	};
}

const EQUATION_STYLE = { fontSize: 18, fontFamily: 'Cambria Math' };

/** The shared session's port over the Vue viewer's operations, read at call time. */
export function createEngineSessionPort(c: Context): EngineSessionPort {
	const selected = () => c.selectedElements.value[0] ?? null;
	const patchSelection = (patch: Partial<PptxElement>) => {
		const element = selected();
		if (!element) {
			return buildingBlockFailure('no-selection');
		}
		c.ops.updateElement(element.id, patch);
		return BUILDING_BLOCK_OK;
	};
	return {
		availabilityInput: () => availabilityInput(c),
		target: () => ({
			deck: c.handler.value,
			slideId: c.activeSlide.value?.id,
			slideIndex: c.activeSlideIndex.value,
		}),
		canvasSize: () => c.canvasSize.value,
		selectedElement: selected,
		slideCount: () => c.slides.value.length,
		activeSlide: () => c.activeSlide.value,
		scopedLayouts: () => scopedLayouts(c),
		addElement: (element) => c.ops.addElement(element),
		insertMediaFile: async (file, sameTarget) => {
			const element = await mediaElementFromFile(file, c.canvasSize.value);
			if (!element) {
				return buildingBlockFailure('decode-failed');
			}
			if (!sameTarget()) {
				return buildingBlockFailure('stale-target');
			}
			c.ops.addElement(element);
			return BUILDING_BLOCK_OK;
		},
		applyHyperlink: (data) => {
			const patch = hyperlinkEditPatch(data);
			return patch.ok ? patchSelection({ actionClick: patch.actionClick }) : patch;
		},
		removeHyperlink: () => {
			const element = selected();
			if (element && !element.actionClick) {
				return buildingBlockFailure('no-hyperlink');
			}
			return patchSelection({ actionClick: undefined });
		},
		insertSmartArt: (layout, items) => c.insertDialogs.onInsertSmartArt(layout, items),
		insertChart: (kind, options) => c.insertion.addChart(kind, options),
		insertEquation: (omml) =>
			c.ops.addElement({
				id: createEditorId('equation'),
				type: 'shape',
				x: 120,
				y: 200,
				width: 400,
				height: 80,
				text: '[Equation]',
				textStyle: EQUATION_STYLE,
				textSegments: [{ text: '[Equation]', style: EQUATION_STYLE, equationXml: omml }],
			} as unknown as PptxElement),
		updateEquation: (omml) => {
			const element = selected();
			if (element) {
				c.ops.updateElement(element.id, {
					textSegments: [{ text: '[Equation]', style: EQUATION_STYLE, equationXml: omml }],
				} as Partial<PptxElement>);
			}
		},
		findReplace: () => findPort(c),
		applyLayout: async (path) => {
			if (!c.handler.value || !c.activeSlide.value) {
				throw new Error(`Layout ${path} was not applied`);
			}
			await c.insertion.applyLayoutToActiveSlide(path);
		},
	};
}

/** The ports of the shared dialog groups over the viewer's own composables. */
export function createEngineGroupPorts(c: Context): BuildingBlockGroupPorts {
	const markDirty = () => {
		c.isDirty.value = true;
	};
	return {
		passwordProtection: {
			input: () => ({
				openPasswordSet: c.password.isPasswordProtected.value,
				recommendation: { kind: c.readOnlyRecommendation.value.kind },
				hasModifyVerifier: c.modifyVerifier.value !== undefined,
				customProperties: c.customProperties.value,
			}),
			customProperties: () => c.customProperties.value,
			applyOpenPassword: (password) =>
				password === null ? c.password.onRemovePassword() : c.password.onSetPassword(password),
			applyCustomProperties: (next) => c.deckActions.updateCustomProperties(next),
			markDirty,
		},
		documentProperties: {
			current: () => ({
				core: c.coreProperties.value ?? {},
				custom: c.customProperties.value,
				app: c.appProperties.value ?? {},
			}),
			applyUpdate: (update) => {
				if (update.core) {
					c.deckActions.updateCoreProperties(update.core);
				}
				if (update.custom) {
					c.deckActions.updateCustomProperties(update.custom);
				}
				if (update.app) {
					c.deckActions.updateAppProperties(update.app);
				}
			},
		},
		fontEmbedding: {
			input: () => ({
				embedFonts: c.fontEmbedding.embedFontsEnabled.value,
				embeddedFamilies: c.fontEmbedding.embeddedFontNames.value,
				usedFamilies: c.fontEmbedding.usedFontFamilies.value,
				descriptor: c.fontEmbedding.fontEmbedding.value,
			}),
			apply: (enabled) => {
				c.fontEmbedding.embedFontsEnabled.value = enabled;
			},
			markDirty,
		},
		digitalSignatures: {
			signed: () => c.signatureWorkflow.hasDigitalSignatures.value,
			count: () => c.signatures.value.length,
			content: () => c.activeContent.value,
			markDirty,
		},
		setUpSlideShow: {
			presentationProperties: () => c.presentationProperties.value,
			slideCount: () => c.slides.value.length,
			customShows: () => c.customShows.value,
			save: (settings) => {
				c.slideShow.onSaveSlideShowSettings(settings);
				markDirty();
			},
		},
		customShows: {
			customShows: () => c.customShows.value,
			slides: () => c.slides.value,
			activeId: () => c.customShowsWiring.activeCustomShowId.value,
			commit: (shows, slides) => c.customShowsWiring.commitCustomShowsDialog(shows, slides),
			start: (id) => c.customShowsWiring.runCustomShow(id),
		},
		shortcuts: {
			open: () => c.showShortcuts.value,
			setOpen: (open) => {
				c.showShortcuts.value = open;
			},
			toggle: () => {
				c.showShortcuts.value = !c.showShortcuts.value;
			},
			translate: (key) => c.t(key),
		},
	};
}
