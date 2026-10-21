/**
 * ExportProgressModal: the progress modal shown while a multi-slide export (PDF, GIF, video)
 * runs. A thin adapter over the shared `pptx-ui-export-progress`, which owns the dialog and
 * progress bar semantics, focus and the look; the host owns cancelling the export.
 */
import type { PptxUiExportProgressElement } from 'ooxml-ui/pptx';
import { useTranslation } from 'react-i18next';

import { useWebControl } from '../hooks/useWebControl';

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

export interface ExportProgressModalProps {
	/** Whether the modal is visible. */
	isOpen: boolean;
	/** Title shown at the top of the modal. */
	title: string;
	/** Current progress (0-100). */
	progress: number;
	/** Optional status message. */
	statusMessage?: string;
	/** Callback to cancel the export. */
	onCancel: () => void;
}

/* ------------------------------------------------------------------ */
/*  Component                                                          */
/* ------------------------------------------------------------------ */

export function ExportProgressModal({
	isOpen,
	title,
	progress,
	statusMessage,
	onCancel,
}: ExportProgressModalProps): React.ReactElement {
	const { t } = useTranslation();
	// Kept mounted while closed, so it can hand focus back when an export ends.
	const ref = useWebControl<PptxUiExportProgressElement>(
		{ open: isOpen, title, progress, status: statusMessage, translate: t },
		{ 'export-cancel': () => onCancel() },
	);
	return <pptx-ui-export-progress ref={ref} />;
}
