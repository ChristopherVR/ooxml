import type {
	InsertChartDialogInsertEvent,
	InsertChartOptions,
	InsertChartKind,
	InsertSmartArtDialogInsertEvent,
	PptxUiInsertChartDialogElement,
	PptxUiInsertSmartArtDialogElement,
} from 'ooxml-ui/pptx';
import type { SmartArtLayout } from 'pptx-viewer-core';
import React, { useEffect, useLayoutEffect, useRef } from 'react';

import { useTranslation } from '../i18n/use-translation';

/**
 * React 18 and 19 adapter for the shared Insert SmartArt and Insert Chart dialogs
 * (`pptx-ui-insert-smartart-dialog`, `pptx-ui-insert-chart-dialog`): the elements own the dialog,
 * the picker and the previews; the binding only maps its open state and inserts the result.
 */
function useDialogEvents<E extends HTMLElement>(
	ref: React.RefObject<E | null>,
	insertEvent: string,
	onInsert: (event: Event) => void,
	onClose: () => void,
): void {
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		host.addEventListener(insertEvent, onInsert);
		host.addEventListener('dialog-close', onClose);
		return () => {
			host.removeEventListener(insertEvent, onInsert);
			host.removeEventListener('dialog-close', onClose);
		};
	}, [ref, insertEvent, onInsert, onClose]);
}

export function InsertSmartArtDialog({
	isOpen,
	onClose,
	onInsert,
}: {
	isOpen: boolean;
	onClose: () => void;
	onInsert: (layout: SmartArtLayout, items: string[]) => void;
}): React.ReactElement {
	const ref = useRef<PptxUiInsertSmartArtDialogElement>(null);
	const { t } = useTranslation();
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.translateLabel = t;
			ref.current.open = isOpen;
		}
	}, [isOpen, t]);
	useDialogEvents(
		ref,
		'insert-smartart',
		(event) => {
			const { layout, items } = (event as InsertSmartArtDialogInsertEvent).detail;
			onInsert(layout, items);
		},
		onClose,
	);
	return <pptx-ui-insert-smartart-dialog ref={ref} />;
}

export function InsertChartDialog({
	isOpen,
	onClose,
	onInsert,
}: {
	isOpen: boolean;
	onClose: () => void;
	onInsert: (kind: InsertChartKind, options: InsertChartOptions) => void;
}): React.ReactElement {
	const ref = useRef<PptxUiInsertChartDialogElement>(null);
	const { t } = useTranslation();
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.translateLabel = t;
			ref.current.open = isOpen;
		}
	}, [isOpen, t]);
	useDialogEvents(
		ref,
		'insert-chart',
		(event) => {
			const { kind, title, grouping } = (event as InsertChartDialogInsertEvent).detail;
			onInsert(kind, { title, grouping });
		},
		onClose,
	);
	return <pptx-ui-insert-chart-dialog ref={ref} />;
}
