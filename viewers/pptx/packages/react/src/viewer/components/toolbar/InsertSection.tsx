import {
	DEFAULT_INSERT_CHART_KIND,
	FREEFORM_TOOL_IDS,
	isDialogAvailable,
	isDrawingToolVisible,
	resolveInsertSlideLayout,
} from 'ooxml-ui/pptx';
import type {
	FreeformToolKind,
	InsertChartKind,
	InsertSlideLayout,
	PptxUiMobileInsertElement,
	PptxUiRibbonInsertElement,
	SlideTemplateId,
	RibbonInsertRequestEvent,
} from 'ooxml-ui/pptx';
import React, { useEffect, useRef, useState } from 'react';
import { useTranslation } from 'react-i18next';

import type { SupportedShapeType } from '../../types';
import { SlideTemplateGalleryDialog } from '../SlideTemplateGalleryDialog';
import { useViewerCustomizationContext } from '../viewer-customization-context';
import { DateTimeFieldDialog } from './DateTimeFieldDialog';

export interface InsertSectionProps {
	canEdit: boolean;
	newShapeType: SupportedShapeType;
	onSetNewShapeType: (type: SupportedShapeType) => void;
	/** The armed Freeform: Shape / Curve tool, or null. */
	activeFreeformTool?: FreeformToolKind | null;
	/** Arm (or, with null, disarm) a Freeform: Shape / Curve tool. */
	onArmFreeformTool?: (tool: FreeformToolKind | null) => void;
	onAddTextBox: () => void;
	onAddShape: () => void;
	onAddTable: (rows?: number, columns?: number) => void;
	onAddChart?: (chartKind: InsertChartKind) => void;
	onAddSmartArt: () => void;
	onAddEquation: () => void;
	onAddActionButton: (shapeType: string) => void;
	onInsertField?: (fieldType: string, value?: string) => void;
	onOpenHeaderFooter?: () => void;
	onOpenImagePicker: () => void;
	onOpenMediaPicker: () => void;
	/** True when something is selected, so a link has a target to attach to. */
	hasSelection: boolean;
	onOpenHyperlinkDialog: () => void;
	/** The layouts New Slide offers, and the active slide's layout. */
	layoutOptions?: readonly InsertSlideLayout[];
	currentLayoutPath?: string;
	onInsertSlideFromLayout?: (path: string, name?: string) => void;
	/** Slide Templates... ends New Slide's menu; without a handler the entry is left out. */
	onInsertSlideFromTemplate?: (templateId: SlideTemplateId) => void;
	/** Deck scheme map so template previews show the deck's theme colours. */
	templateScheme?: Record<string, string>;
	/** Comment opens the comments pane, as Review > New Comment does. */
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	/** Draw the phone Insert sheet (`pptx-ui-mobile-insert`) instead of the ribbon group row. */
	phone?: boolean;
	/** Phone only: called after a request is handled, so the sheet can close. */
	onRequestHandled?: () => void;
}

/**
 * Thin adapter over `pptx-ui-ribbon-insert`. Document mutation, the file/SmartArt/
 * equation/hyperlink/header-footer dialogs and the Date/Time picker stay native.
 */
export function InsertSection(p: InsertSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const customization = useViewerCustomizationContext();
	const ref = useRef<PptxUiRibbonInsertElement | PptxUiMobileInsertElement>(null);
	const [datePickerOpen, setDatePickerOpen] = useState(false);
	const [templatesOpen, setTemplatesOpen] = useState(false);
	const templatesAvailable =
		Boolean(p.onInsertSlideFromTemplate) && isDialogAvailable(customization, 'slideTemplates');
	const [chartKind, setChartKind] = useState<InsertChartKind>(DEFAULT_INSERT_CHART_KIND);
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		host.state = {
			editable: p.canEdit,
			hasSelection: p.hasSelection,
			shapeType: p.newShapeType,
			chartKind,
			activeFreeformTool: p.activeFreeformTool ?? null,
			freeformTools: p.onArmFreeformTool
				? FREEFORM_TOOL_IDS.filter((tool) => isDrawingToolVisible(customization, tool))
				: [],
			chartAvailable: Boolean(p.onAddChart),
			fieldAvailable: Boolean(p.onInsertField),
			headerFooterAvailable: Boolean(p.onOpenHeaderFooter),
			newSlideAvailable: Boolean(p.onInsertSlideFromLayout && p.layoutOptions?.length),
			commentAvailable: Boolean(p.onToggleComments),
			slideTemplatesAvailable: templatesAvailable,
			layouts: p.layoutOptions ?? [],
			...(p.currentLayoutPath ? { currentLayoutPath: p.currentLayoutPath } : {}),
			translate: t,
		};
	}, [p, chartKind, customization, t, templatesAvailable]);
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const request = (event: Event) => {
			const intent = (event as RibbonInsertRequestEvent).detail;
			switch (intent.kind) {
				case 'command': {
					const commands = {
						textBox: p.onAddTextBox,
						table: p.onAddTable,
						image: p.onOpenImagePicker,
						media: p.onOpenMediaPicker,
						smartArt: p.onAddSmartArt,
						equation: p.onAddEquation,
						link: p.onOpenHyperlinkDialog,
						headerFooter: p.onOpenHeaderFooter,
						// Opening an already open pane would toggle it shut.
						comment: p.isCommentsPanelOpen ? undefined : p.onToggleComments,
					};
					commands[intent.value]?.();
					break;
				}
				case 'table':
					p.onAddTable(intent.rows, intent.columns);
					break;
				case 'newSlide': {
					const layout = resolveInsertSlideLayout(
						p.layoutOptions ?? [],
						intent.value,
						p.currentLayoutPath,
					);
					if (layout) {
						p.onInsertSlideFromLayout?.(layout.path, layout.name);
					}
					break;
				}
				case 'slideTemplates':
					setTemplatesOpen(true);
					break;
				case 'shapeType':
					p.onSetNewShapeType(intent.value as SupportedShapeType);
					break;
				case 'shape':
					p.onAddShape();
					break;
				case 'chartType':
					setChartKind(intent.value as InsertChartKind);
					break;
				case 'chart':
					p.onAddChart?.(intent.value as InsertChartKind);
					break;
				case 'freeform':
					p.onArmFreeformTool?.(intent.value as FreeformToolKind | null);
					break;
				case 'actionButton':
					p.onAddActionButton(intent.value);
					break;
				case 'field':
					if (intent.value === 'datetime') {
						setDatePickerOpen(true);
					} else {
						p.onInsertField?.(intent.value);
					}
			}
			p.onRequestHandled?.();
		};
		host.addEventListener('insert-request', request);
		return () => host.removeEventListener('insert-request', request);
	}, [p]);
	return (
		<>
			{p.phone ? (
				<pptx-ui-mobile-insert ref={ref as React.Ref<PptxUiMobileInsertElement>} />
			) : (
				<pptx-ui-ribbon-insert ref={ref as React.Ref<PptxUiRibbonInsertElement>} />
			)}
			{templatesAvailable && (
				<SlideTemplateGalleryDialog
					isOpen={templatesOpen}
					onClose={() => setTemplatesOpen(false)}
					onInsert={(templateId) => p.onInsertSlideFromTemplate?.(templateId)}
					scheme={p.templateScheme}
				/>
			)}
			{datePickerOpen && p.onInsertField && (
				<DateTimeFieldDialog
					onClose={() => setDatePickerOpen(false)}
					onInsert={(formatted) => {
						p.onInsertField?.('datetime', formatted);
						setDatePickerOpen(false);
					}}
				/>
			)}
		</>
	);
}
