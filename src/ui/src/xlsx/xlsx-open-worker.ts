/// <reference lib="webworker" />
// Runs the calculation a workbook asks for on open off the page: reads the file the page is also
// reading, calculates it within the time limit, and posts back only the cells the calculation
// wrote. `workbook-opener.ts` drives it, and ends it should a calculation not stop in time.
import { calculateOnOpenCells } from 'ooxml-core/xlsx';
import { loadWorkbook } from 'ooxml-core/xlsx/load';
import type { CalculateRequest, CalculateResponse } from './workbook-opener';

const worker = self as unknown as DedicatedWorkerGlobalScope;
const reply = (message: CalculateResponse, transfer: Transferable[] = []): void =>
	worker.postMessage(message, transfer);

worker.onmessage = async (event: MessageEvent<CalculateRequest>) => {
	if (event.origin && event.origin !== self.location.origin) return;
	const { bytes, fileName, password, timeLimitMs } = event.data;
	try {
		const workbook = await loadWorkbook(bytes, {
			...(fileName === undefined ? {} : { fileName }),
			...(password === undefined ? {} : { password }),
		});
		reply({ kind: 'read' });
		const { calculation, cells } = calculateOnOpenCells(workbook, { timeLimitMs });
		if (calculation.stopped) return reply({ kind: 'stopped', reason: calculation.stopped });
		if (!cells) return reply({ kind: 'skipped' });
		reply(
			{ kind: 'done', cells },
			cells.sheets.flatMap((sheet) => [
				sheet.cells.buffer,
				sheet.anchors.buffer,
				sheet.released.buffer,
			]),
		);
	} catch (cause) {
		reply({ kind: 'error', message: cause instanceof Error ? cause.message : String(cause) });
	}
};

// Told first, so an error before it means the worker could not start rather than a failed run.
reply({ kind: 'ready' });
