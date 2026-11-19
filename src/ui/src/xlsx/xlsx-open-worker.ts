/// <reference lib="webworker" />
// Opens a workbook off the page: reads the file, then runs the calculation it asks for on open, so
// neither a large file nor a slow formula holds up the editor. `workbook-opener.ts` drives it and
// ends it when the calculation runs past its time limit.
import { calculateOnOpen } from 'ooxml-core/xlsx';
import { loadWorkbook } from 'ooxml-core/xlsx/load';
import type { OpenRequest, OpenResponse } from './workbook-opener';

const worker = self as unknown as DedicatedWorkerGlobalScope;
const reply = (message: OpenResponse): void => worker.postMessage(message);

worker.onmessage = async (event: MessageEvent<OpenRequest>) => {
	if (event.origin && event.origin !== self.location.origin) return;
	const { bytes, fileName, password, calculate } = event.data;
	try {
		const workbook = await loadWorkbook(bytes, {
			...(fileName === undefined ? {} : { fileName }),
			...(password === undefined ? {} : { password }),
		});
		reply({ kind: 'read' });
		let calculated = false;
		try {
			calculated = calculate && calculateOnOpen(workbook);
		} catch (error) {
			// As on the page: the workbook still opens, with the values it was saved with.
			console.error(error);
		}
		reply({ kind: 'done', workbook, calculated });
	} catch (cause) {
		const code = (cause as { code?: unknown } | null)?.code;
		reply({
			kind: 'error',
			message: cause instanceof Error ? cause.message : String(cause),
			...(typeof code === 'string' ? { code } : {}),
		});
	}
};

// Told first, so an error before it means the worker could not start rather than a failed open.
reply({ kind: 'ready' });
