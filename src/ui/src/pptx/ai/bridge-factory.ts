/**
 * `createPptxAiBridge`: the one {@link PptxAiBridge} implementation.
 *
 * Every binding used to carry its own 200 to 250 line copy of the same bridge
 * (deck metadata, the slide write choke points, the whole-deck read and the
 * field-by-field `applyDeckData` fan-out), and the copies had begun to drift
 * (titles, clearing document properties, which binding reported core property
 * changes to the host). The logic lives here now; a binding supplies a
 * {@link PptxAiBridgePort}: live getters and the setters that route each write
 * through its own undo history.
 *
 * This module has no dependency on the optional `ai` SDK, so a binding can
 * build the bridge eagerly and only load the SDK when the panel opens.
 *
 * @module ai/bridge-factory
 */

import type {
	ParsedTableStyleMap,
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxData,
	PptxHandler,
	PptxPresentationProperties,
	PptxSection,
	PptxSlide,
	PptxTagCollection,
	PptxTheme,
	PptxViewProperties,
} from 'ooxml-core/pptx';

import type {
	PptxAiBridge,
	PptxAiDataUpdater,
	PptxAiDeckMeta,
	PptxAiElementUpdate,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
	PptxAiSlidesUpdater,
} from './bridge';
import { deckDataFieldChanged } from './deck-data-diff';
import { applyElementUpdate } from './tools/mutations';

/** Canvas size in CSS pixels. */
export interface PptxAiCanvasSize {
	width: number;
	height: number;
}

/** What a binding supplies so {@link createPptxAiBridge} can drive its editor. */
export interface PptxAiBridgePort {
	// -- Reads (always the live value) ------------------------------------
	getSlides(): readonly PptxSlide[];
	getActiveSlideIndex(): number;
	getCanvasSize(): PptxAiCanvasSize;
	getTheme(): PptxTheme | undefined;
	getHandler(): PptxHandler | null | undefined;
	/** Display name of the document, used as the deck title when present. */
	getFileName?(): string | undefined;
	getSections(): readonly PptxSection[] | undefined;
	getPresentationProperties(): PptxPresentationProperties | undefined;
	getCustomProperties(): readonly PptxCustomProperty[] | undefined;
	getCoreProperties(): PptxCoreProperties | undefined;
	getAppProperties(): PptxAppProperties | undefined;
	getViewProperties(): PptxViewProperties | undefined;
	getTableStyleMap(): ParsedTableStyleMap | undefined;
	getTableStylesDefaultId(): string | undefined;
	getTagCollections(): readonly PptxTagCollection[] | undefined;

	// -- Navigation --------------------------------------------------------
	goToSlide(index: number): void;
	/** Select elements on a slide; the port navigates first when needed. */
	selectElements(slideIndex: number, elementIds: string[]): void;

	// -- Writes ------------------------------------------------------------
	/**
	 * Commit a slides edit as ONE undoable history entry. Call `update` with the
	 * current slides and commit what it returns; a binding with functional state
	 * updates passes it straight through, so two edits in one tick never read
	 * stale slides. The bridge clones before an AI updater mutates anything.
	 */
	updateSlides(update: (current: readonly PptxSlide[]) => PptxSlide[], label: string): void;
	/** Apply partial theme updates (colour / font scheme) and re-render. */
	applyTheme(updates: Partial<PptxTheme>): void;
	setCanvasSize(size: PptxAiCanvasSize): void;
	setSections(sections: PptxSection[]): void;
	setPresentationProperties(properties: PptxPresentationProperties): void;
	/** Core, app and custom document properties, committed as one unit. */
	setDocumentProperties(
		core: PptxCoreProperties,
		app: PptxAppProperties,
		custom: PptxCustomProperty[],
	): void;
	/** View preferences; not undo-tracked, like the manual View toggles. */
	setViewProperties(properties: PptxViewProperties | undefined): void;
	setTableStyleMap(map: ParsedTableStyleMap): void;
	setTableStylesDefaultId(id: string | undefined): void;
	setTagCollections(tags: PptxTagCollection[]): void;
	/** Called once after `applyDeckData` changed anything (mark the deck dirty). */
	afterDeckUpdate?(): void;

	// -- Optional host hooks -------------------------------------------------
	/** The assistant's focus (picks, pin, live selection); see the panel controller. */
	getFocusedTargets?(): PptxAiFocusedTarget[];
	/** The host's `properties-change`: called when an AI edit changed core properties. */
	onCorePropertiesChanged?(properties: PptxCoreProperties): void;
	/** Host notification sink (status line, toast, console). */
	notify?(message: string, level?: PptxAiNotifyLevel): void;
}

const TITLE_MAX_LENGTH = 120;

/** First non-empty text on the first slide, the fallback deck title. */
export function pptxAiDeckTitle(slides: readonly PptxSlide[]): string | undefined {
	for (const element of slides[0]?.elements ?? []) {
		const text = 'text' in element && typeof element.text === 'string' ? element.text.trim() : '';
		if (text) {
			return text.slice(0, TITLE_MAX_LENGTH);
		}
	}
	return undefined;
}

function copy<T>(list: readonly T[] | undefined): T[] | undefined {
	return list ? [...list] : undefined;
}

/** Reconstruct the presentation-level {@link PptxData} the deck tools read and write. */
function readDeckData(port: PptxAiBridgePort): PptxData {
	const canvas = port.getCanvasSize();
	return {
		slides: [...port.getSlides()],
		width: canvas.width,
		height: canvas.height,
		theme: port.getTheme(),
		sections: copy(port.getSections()),
		presentationProperties: port.getPresentationProperties(),
		customProperties: copy(port.getCustomProperties()),
		coreProperties: port.getCoreProperties(),
		appProperties: port.getAppProperties(),
		viewProperties: port.getViewProperties(),
		tableStyleMap: port.getTableStyleMap(),
		tableStylesDefaultId: port.getTableStylesDefaultId(),
		tags: copy(port.getTagCollections()),
	} satisfies Partial<PptxData> as PptxData;
}

/** Fan the changed document-property fields out as one unit. */
function commitDocumentProperties(
	port: PptxAiBridgePort,
	before: PptxData,
	after: PptxData,
): boolean {
	const differs = deckDataFieldChanged;
	const core = after.coreProperties ?? before.coreProperties;
	const app = after.appProperties ?? before.appProperties;
	const custom = after.customProperties ?? before.customProperties ?? [];
	const coreChanged = differs(before.coreProperties, core);
	if (
		!coreChanged &&
		!differs(before.appProperties, app) &&
		!differs(before.customProperties, custom)
	) {
		return false;
	}
	port.setDocumentProperties(core ?? {}, app ?? {}, custom);
	if (coreChanged && core) {
		port.onCorePropertiesChanged?.(core);
	}
	return true;
}

/**
 * Fan the changed top-level deck fields out to the port's setters, each
 * through its own editor path. Returns true when anything changed.
 */
function commitDeckData(
	port: PptxAiBridgePort,
	before: PptxData,
	after: PptxData,
	label: string,
): boolean {
	const differs = deckDataFieldChanged;
	let changed = false;
	if (differs(before.slides, after.slides)) {
		port.updateSlides(() => after.slides, label);
		changed = true;
	}
	if (before.width !== after.width || before.height !== after.height) {
		port.setCanvasSize({ width: after.width, height: after.height });
		changed = true;
	}
	const sections = after.sections ?? before.sections ?? [];
	if (differs(before.sections, sections)) {
		port.setSections(sections);
		changed = true;
	}
	const presentation = after.presentationProperties ?? before.presentationProperties ?? {};
	if (differs(before.presentationProperties, presentation)) {
		port.setPresentationProperties(presentation);
		changed = true;
	}
	changed = commitDocumentProperties(port, before, after) || changed;
	if (differs(before.viewProperties, after.viewProperties)) {
		port.setViewProperties(after.viewProperties);
		changed = true;
	}
	if (differs(before.tableStyleMap, after.tableStyleMap)) {
		port.setTableStyleMap(after.tableStyleMap ?? {});
		changed = true;
	}
	if (differs(before.tableStylesDefaultId, after.tableStylesDefaultId)) {
		port.setTableStylesDefaultId(after.tableStylesDefaultId);
		changed = true;
	}
	const tags = after.tags ?? before.tags ?? [];
	if (differs(before.tags, tags)) {
		port.setTagCollections(tags);
		changed = true;
	}
	return changed;
}

/** Build the bridge that exposes a live editor (through `port`) to the AI core. */
export function createPptxAiBridge(port: PptxAiBridgePort): PptxAiBridge {
	const applySlidesUpdate = (updater: PptxAiSlidesUpdater, label: string): void => {
		port.updateSlides((current) => updater(structuredClone([...current])), label);
	};

	const bridge: PptxAiBridge = {
		getDeckMeta(): PptxAiDeckMeta {
			const slides = port.getSlides();
			const canvas = port.getCanvasSize();
			return {
				slideCount: slides.length,
				activeSlideIndex: port.getActiveSlideIndex(),
				title: port.getFileName?.() ?? pptxAiDeckTitle(slides),
				width: canvas.width,
				height: canvas.height,
			};
		},
		getSlides: () => [...port.getSlides()],
		getActiveSlideIndex: () => port.getActiveSlideIndex(),
		getTheme: () => port.getTheme(),
		getHandler: () => port.getHandler() ?? undefined,
		goToSlide: (index) => port.goToSlide(index),
		selectElements: (slideIndex, elementIds) => port.selectElements(slideIndex, elementIds),
		applySlidesUpdate,
		updateElement(slideIndex: number, elementId: string, updates: PptxAiElementUpdate) {
			applySlidesUpdate((slides) => {
				const element = slides[slideIndex]?.elements.find((el) => el.id === elementId);
				if (element) {
					applyElementUpdate(element, updates);
				}
				return slides;
			}, `Update element ${elementId}`);
		},
		applyTheme: (updates) => port.applyTheme(updates),
		getDeckData: () => readDeckData(port),
		applyDeckData(updater: PptxAiDataUpdater, label: string) {
			const before = readDeckData(port);
			const after = updater(structuredClone(before));
			if (commitDeckData(port, before, after, label)) {
				port.afterDeckUpdate?.();
			}
		},
	};
	if (port.getFocusedTargets) {
		bridge.getFocusedTargets = () => port.getFocusedTargets?.() ?? [];
	}
	if (port.notify) {
		bridge.notify = (message, level) => port.notify?.(message, level);
	}
	return bridge;
}
