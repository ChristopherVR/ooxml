import React from 'react';
import { useTranslation } from 'react-i18next';

import { InsertSection } from '../toolbar/InsertSection';
import type { ToolbarProps } from '../toolbar/toolbar-types';
import { MobileSheet } from './MobileSheet';

/**
 * The bottom bar's Insert sheet: the shared `pptx-ui-mobile-insert` tile grid, routed through
 * the same `InsertSection` handling as the ribbon, so a pick runs the identical command.
 */
export function MobileInsertSheet(p: ToolbarProps): React.ReactElement {
	const { t } = useTranslation();
	const close = () => p.onCloseMobileInsert?.();
	return (
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
					onToggleComments={p.onToggleComments}
					isCommentsPanelOpen={p.isCommentsPanelOpen}
				/>
			</div>
		</MobileSheet>
	);
}
