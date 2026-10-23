import { isDialogAvailable } from 'ooxml-ui/pptx';
import React, { useMemo, useState } from 'react';
import { useTranslation } from 'react-i18next';

import { SlideTemplateGalleryDialog } from '../SlideTemplateGalleryDialog';
import { DateTimeFieldDialog } from '../toolbar/DateTimeFieldDialog';
import { InsertSection } from '../toolbar/InsertSection';
import type { ToolbarProps } from '../toolbar/toolbar-types';
import { useViewerCustomizationContext } from '../viewer-customization-context';
import { MobileSheet } from './MobileSheet';

/**
 * The bottom bar's Insert sheet: the shared `pptx-ui-mobile-insert` tile grid, routed through
 * the same `InsertSection` handling as the ribbon, so a pick runs the identical command.
 */
export function MobileInsertSheet(p: ToolbarProps): React.ReactElement {
	const { t } = useTranslation();
	const close = () => p.onCloseMobileInsert?.();
	const customization = useViewerCustomizationContext();
	// These dialogs are opened from the sheet but must outlive it: the sheet unmounts its body
	// the moment a pick closes it, which would take a dialog owned by the section along.
	const [dateOpen, setDateOpen] = useState(false);
	const [templatesOpen, setTemplatesOpen] = useState(false);
	const dialogHost = useMemo(
		() => ({
			openDateTime: () => setDateOpen(true),
			openSlideTemplates: () => setTemplatesOpen(true),
		}),
		[],
	);
	return (
		<>
			<MobileSheet
				open={Boolean(p.isMobileInsertOpen) && !p.isThemeEditorOpen}
				onClose={close}
				autoHeight
				title={t('pptx.mobileBar.insert')}
			>
				<div className='p-3'>
					<InsertSection
						phone
						onRequestHandled={close}
						canEdit={p.canEdit}
						newShapeType={p.newShapeType}
						onSetNewShapeType={p.onSetNewShapeType}
						onAddTextBox={p.onAddTextBox}
						onAddShape={p.onAddShape}
						onAddTable={p.onAddTable}
						onAddChart={p.onAddChart}
						onOpenChartDialog={p.onOpenChartDialog}
						onAddSmartArt={p.onAddSmartArt}
						onAddEquation={p.onAddEquation}
						onAddActionButton={p.onAddActionButton}
						onInsertField={p.onInsertField}
						onOpenHeaderFooter={p.onOpenHeaderFooter}
						onOpenImagePicker={p.onOpenImagePicker}
						onOpenMediaPicker={p.onOpenMediaPicker}
						hasSelection={Boolean(p.selectedElement)}
						onOpenHyperlinkDialog={p.onOpenHyperlinkDialog}
						layoutOptions={p.layoutOptions}
						currentLayoutPath={p.currentLayoutPath}
						onInsertSlideFromLayout={p.onInsertSlideFromLayout}
						onInsertSlideFromTemplate={p.onInsertSlideFromTemplate}
						templateScheme={p.templateScheme}
						onToggleComments={p.onToggleComments}
						isCommentsPanelOpen={p.isCommentsPanelOpen}
						dialogHost={dialogHost}
					/>
				</div>
			</MobileSheet>
			{p.onInsertSlideFromTemplate && isDialogAvailable(customization, 'slideTemplates') && (
				<SlideTemplateGalleryDialog
					isOpen={templatesOpen}
					onClose={() => setTemplatesOpen(false)}
					onInsert={(templateId) => p.onInsertSlideFromTemplate?.(templateId)}
					scheme={p.templateScheme}
				/>
			)}
			{dateOpen && p.onInsertField && (
				<DateTimeFieldDialog
					onClose={() => setDateOpen(false)}
					onInsert={(formatted) => {
						p.onInsertField?.('datetime', formatted);
						setDateOpen(false);
					}}
				/>
			)}
		</>
	);
}
