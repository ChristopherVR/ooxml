import {
	createFormatPaneController,
	formatSelectionElements,
	tableStyleAssignmentUpdate,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { createIcon } from '../icons';
import type { TableStyleEditorDeps } from '../table-style-editor';
import { createAccessibilitySection } from './accessibility-section';
import { createActionSection } from './action-section';
import { createAnimationPanel } from './animation-panel';
import { createChartSection } from './chart-section';
import { createDeckPanel } from './deck-panel';
import { createElementsTab } from './elements-tab';
import { createFillSection } from './fill-section';
import { createFormatSections } from './format-sections';
import { createGroupInfoSection } from './group-info-section';
import { createImageSection } from './image-section';
import { createMediaSection } from './media-section';
import { createOlePropertiesSection } from './ole-properties-section';
import { createPositionSection } from './position-section';
import { createQuickStylesGallery } from './quick-styles-gallery';
import { createSmartArtSection } from './smartart-section';
import { createTableDataGrid } from './table-data-grid';
import { createTableSection } from './table-section';
import { createText3DSection } from './text-3d-section';
import { createTextSection } from './text-section';
import type { Inspector, InspectorHandlers } from './types';

type InspectorTabId = 'elements' | 'properties';

/**
 * The property inspector: a right-hand panel with React's three-tab strip
 * (Elements | Properties | Comments). The Properties tab shows the
 * element-type-aware sections when something is selected and a deck-level
 * panel (presentation / slide size / document) otherwise; Elements lists the
 * slide's layer order; Comments carries the slide's review comments. Vanilla
 * counterpart of React's `InspectorPane`, scoped to the highest-traffic
 * panels. All labels from the shared `pptx.*` dictionary.
 */
export function createInspector(
	doc: Document,
	t: Translator,
	handlers: InspectorHandlers,
): Inspector {
	const el = createEl(doc, 'aside', 'pptxv-inspector');
	el.setAttribute('data-pptx-inspector', '');
	el.dataset.pptxChrome = 'inspector';
	el.setAttribute('aria-label', t('pptx.inspector.properties'));

	// -- Tab strip (React's InspectorPaneHeader) ------------------------------
	const header = createEl(doc, 'div', 'pptxv-inspector-tabs');
	header.setAttribute('role', 'tablist');
	header.dataset.pptxChrome = 'inspector-tabs';
	const headerWrap = createEl(doc, 'div');
	headerWrap.dataset.pptxChrome = 'inspector-header';
	headerWrap.appendChild(header);
	el.appendChild(headerWrap);
	const tabButtons = new Map<InspectorTabId, HTMLButtonElement>();
	const tabDefs: Array<{ id: InspectorTabId; label: string }> = [
		{ id: 'elements', label: t('pptx.documentProperties.statistics.elements') },
		{ id: 'properties', label: t('pptx.inspector.properties') },
	];
	for (const tab of tabDefs) {
		const btn = createEl(doc, 'button', 'pptxv-inspector-tab');
		btn.type = 'button';
		btn.setAttribute('role', 'tab');
		btn.appendChild(createIcon(doc, tab.id === 'elements' ? 'layers' : 'settings'));
		btn.appendChild(doc.createTextNode(tab.label));
		btn.addEventListener('click', () => setActiveTab(tab.id));
		header.appendChild(btn);
		tabButtons.set(tab.id, btn);
	}

	// -- Tab panes ------------------------------------------------------------
	const elementsTab = createElementsTab(doc, t, (id) => handlers.selectElement(id));
	elementsTab.el.classList.add('pptxv-inspector-body');
	el.appendChild(elementsTab.el);

	const body = createEl(doc, 'div', 'pptxv-inspector-body');
	body.dataset.pptxChrome = 'inspector-body';
	el.appendChild(body);

	// Docked per-element Animation panel: always at the bottom of the pane,
	// visible (on any tab) whenever an element is selected, matching React's
	// `InspectorPane` placement below the tab body.
	const animationPanel = createAnimationPanel(doc, t, handlers);
	el.appendChild(animationPanel.el);
	const nativeTimeline = doc.createElement('pptx-ui-animation-effect-rows');
	nativeTimeline.setAttribute('all', '');
	el.appendChild(nativeTimeline);

	let activeTab: InspectorTabId = 'properties';
	const setActiveTab = (tab: InspectorTabId): void => {
		activeTab = tab;
		for (const [id, btn] of tabButtons) {
			btn.classList.toggle('is-active', id === tab);
			btn.setAttribute('aria-selected', String(id === tab));
		}
		elementsTab.el.hidden = tab !== 'elements';
		body.hidden = tab !== 'properties';
	};

	// -- Properties tab content ----------------------------------------------
	// A Format pane: every panel below is built into the shared section
	// `FORMAT_PANEL_SECTIONS` assigns it (see `format-sections.ts`).
	const pane = handlers.formatPane ?? createFormatPaneController();
	const format = createFormatSections(doc, t, pane, body, () => setActiveTab('properties'));
	body.appendChild(format.el);
	const into = format.into;

	// Deck-level (no selection) cards, React's default inspector view: the Slide section.
	const deckPanel = createDeckPanel(doc, t, handlers);
	format.host('slide').appendChild(deckPanel.el);

	const position = createPositionSection(
		doc,
		t,
		into('sizePosition'),
		handlers.setGeometry,
		handlers.toggleElementLock,
		into('arrange')(t('pptx.inspector.element')),
	);
	const groupInfo = createGroupInfoSection(doc, t, into('arrange'));
	const oleProperties = createOlePropertiesSection(doc, t, into('object'), handlers);
	const fill = createFillSection(doc, t, into('fill'), handlers, {
		outline: into('outline')(t('pptx.inspector.line')),
		effects: into('effects')(t('pptx.inspector.effects')),
		shapeType: into('arrange')(t('pptx.shape.type')),
	});
	const quickStyles = createQuickStylesGallery(doc, t, into('fill'), handlers);
	const text = createTextSection(
		doc,
		t,
		into('font'),
		handlers,
		into('paragraph')(t('pptx.formatPane.paragraph')),
	);
	const text3d = createText3DSection(doc, t, into('font'), handlers);
	const image = createImageSection(doc, t, into('picture'), handlers);
	const accessibility = createAccessibilitySection(doc, t, into('altText'), handlers);
	// The cell-text spreadsheet sits ABOVE the table's styling section, matching
	// React's inspector order. It builds its own <section> (rather than using the
	// `section()` factory) because it needs an aria-labelled landmark, so it is
	// appended to the Table section by hand at exactly this point.
	const tableDataGrid = createTableDataGrid(doc, t, handlers);
	format.host('table').appendChild(tableDataGrid.el);
	// `handlers` already carries everything `TableStyleEditorDeps` needs
	// (getTableStyleMap/getThemeColorMap/updateTableStyleMap/deleteTableStyle),
	// so there is no need to thread a second deps object down from `ChromeOptions`.
	const styleEditorDeps: TableStyleEditorDeps = {
		getTableStyleMap: () => handlers.getTableStyleMap(),
		getThemeColorMap: () => handlers.getThemeColorMap(),
		onStyleMapChange: (map) => handlers.updateTableStyleMap(map),
		onDeleteStyle: (id) => handlers.deleteTableStyle(id),
		onAssignStyle: (styleId) => handlers.setTableOptions(tableStyleAssignmentUpdate(styleId)),
	};
	const table = createTableSection(doc, t, into('table'), handlers, styleEditorDeps);
	const smartArt = createSmartArtSection(doc, t, into('smartArt'), handlers);
	const action = createActionSection(doc, t, into('action'), handlers);
	const chart = createChartSection(doc, t, into('chart'), handlers);
	const media = createMediaSection(doc, t, into('media'), handlers);
	const sections = [
		position,
		groupInfo,
		oleProperties,
		fill,
		quickStyles,
		text,
		text3d,
		image,
		accessibility,
		tableDataGrid,
		table,
		smartArt,
		action,
		chart,
		media,
	];

	setActiveTab(activeTab);

	return {
		el,
		update(state) {
			deckPanel.setVisible(!state.hasSelection);
			for (const s of sections) {
				s.update(state);
			}
		},
		updateDeck(state) {
			nativeTimeline.hidden = Boolean(state.selectedElementId);
			action.setSlideCount(state.slideCount);
			action.setCustomShows(state.customShows);
			elementsTab.update(state);
			deckPanel.update(state);
			format.layout(
				formatSelectionElements(state.selectedElement, state.selectedIds, state.elements),
			);
			animationPanel.update({
				editable: state.editable,
				selectedElementId: state.selectedElementId,
				elements: state.elements,
				animations: state.activeSlide?.animations ?? [],
				animationTimelineAnchors: state.activeSlide?.animationTimelineAnchors ?? [],
			});
		},
		setEditable(editable) {
			el.hidden = !editable;
			tableDataGrid.setEditable(editable);
			chart.setEditable(editable);
		},
	};
}
