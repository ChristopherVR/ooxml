import type { TableCellCaret, TableCellCss, TableCellEditEnd } from 'ooxml-ui/pptx';
import { openTableCellTextEditor, tableCellEditorSegments } from 'ooxml-ui/pptx';
import type { PptxTableCell } from 'pptx-viewer-core';
import React, { useEffect, useRef } from 'react';

/**
 * Rich inline editor for a table cell in editing mode: mounts the shared
 * `openTableCellTextEditor` (the same editor every binding uses) so the
 * cell's runs keep their formatting, Enter adds a paragraph and Tab, the
 * arrows at the text's edge and Escape leave the cell. The binding only maps
 * the reported `TableCellEditEnd` (see `resolveTableCellEditEnd`).
 */
export function TableCellInput({
	cell,
	cellCss,
	caret,
	rtl,
	onEnd,
}: {
	cell: PptxTableCell | undefined;
	cellCss?: TableCellCss;
	caret?: TableCellCaret;
	rtl?: boolean;
	onEnd: (end: TableCellEditEnd) => void;
}) {
	const hostRef = useRef<HTMLDivElement>(null);
	const onEndRef = useRef(onEnd);
	onEndRef.current = onEnd;

	// Mounted once per edited cell. Unmounting (tapping away on mobile
	// deselects the element before the browser fires blur) still commits:
	// `destroy` reports a `detach` end when the editor was open.
	useEffect(() => {
		const host = hostRef.current;
		if (!host) {
			return;
		}
		const session = openTableCellTextEditor({
			host,
			segments: tableCellEditorSegments(cell),
			...(cellCss ? { cellCss } : {}),
			...(caret ? { caret } : {}),
			rtl: rtl === true,
			onEnd: (end) => onEndRef.current(end),
		});
		return () => session.destroy();
		// The seed is read once: the editor owns the text while it is open.
		// oxlint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	return (
		<div
			ref={hostRef}
			className='w-full'
			// Touch surfaces drive canvas drag/marquee through onPointerDown (see
			// useCanvasEventHandlers.handleStagePointerDown). Without stopping it
			// here, tapping inside the cell editor to reposition the caret would
			// bubble to the stage, steal pointer capture and blur the editor.
			onPointerDown={(e) => e.stopPropagation()}
			onMouseDown={(e) => e.stopPropagation()}
			onClick={(e) => e.stopPropagation()}
			onDoubleClick={(e) => e.stopPropagation()}
		/>
	);
}
