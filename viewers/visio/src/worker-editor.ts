import type { VisioDocument, VisioTextEdit } from 'ooxml-core/visio';
import type { EditDiagnostic } from './document-history.js';
import { MAX_INPUT_BYTES } from './scene-validation.js';
export interface EditWorkerRequest {
	bytes: ArrayBuffer;
	edits: readonly VisioTextEdit[];
}
export interface EditTransactionResult {
	bytes: Uint8Array;
	document: VisioDocument;
	changedParts: readonly string[];
	diagnostics: readonly EditDiagnostic[];
}
export type CancellableEditor = ((
	bytes: Uint8Array,
	edits: readonly VisioTextEdit[],
) => Promise<EditTransactionResult>) & { cancel?: () => void };
export interface EditWorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: EditWorkerRequest, transfer: Transferable[]): void;
	terminate(): void;
}
export function createWorkerEditor(
	factory?: () => EditWorkerLike,
	timeoutMs = 20_000,
): CancellableEditor {
	const create =
		factory ??
		(() => {
			if (typeof Worker === 'undefined')
				throw new Error('Editing requires browser worker support.');
			return new Worker(new URL('./edit-worker.js', import.meta.url), { type: 'module' });
		});
	let cancel: (() => void) | undefined;
	const edit: CancellableEditor = (bytes, edits) => {
		cancel?.();
		if (bytes.byteLength > MAX_INPUT_BYTES)
			return Promise.reject(new Error('This viewer accepts files up to 32 MiB.'));
		// Bound UI-thread cloning before the core repeats authoritative command validation.
		if (!Array.isArray(edits) || edits.length > 1000)
			return Promise.reject(
				new Error('At most 1000 text replacements are accepted per operation.'),
			);
		let characters = 0;
		const commands: VisioTextEdit[] = [];
		for (const command of edits) {
			if (
				!command ||
				command.type !== 'replace-plain-text' ||
				typeof command.pageId !== 'string' ||
				!command.pageId ||
				command.pageId.length > 256 ||
				typeof command.shapeId !== 'string' ||
				!command.shapeId ||
				command.shapeId.length > 256 ||
				typeof command.text !== 'string'
			)
				return Promise.reject(new Error('Invalid plain-text edit command.'));
			characters += command.text.length;
			if (characters > 1_000_000)
				return Promise.reject(
					new Error('Replacement text exceeds the one-million-character limit.'),
				);
			// Never clone arbitrary extra properties supplied by an imperative API caller.
			commands.push({
				type: 'replace-plain-text',
				pageId: command.pageId,
				shapeId: command.shapeId,
				text: command.text,
			});
		}
		const copy = Uint8Array.from(bytes);
		return new Promise((resolve, reject) => {
			let worker: EditWorkerLike;
			try {
				worker = create();
			} catch (error) {
				reject(error);
				return;
			}
			let settled = false;
			const finish = (error?: Error, result?: EditTransactionResult) => {
				if (settled) return;
				settled = true;
				clearTimeout(timer);
				cancel = undefined;
				try {
					worker.onmessage = null;
					worker.onerror = null;
					worker.terminate();
				} catch (cause) {
					error ??= cause instanceof Error ? cause : new Error(String(cause));
				}
				if (error) reject(error);
				else resolve(result!);
			};
			const timer = setTimeout(
				() => finish(new Error('The diagram exceeded the isolated editing time limit.')),
				timeoutMs,
			);
			cancel = () => finish(new DOMException('The diagram edit was cancelled.', 'AbortError'));
			worker.onerror = (event) =>
				finish(new Error(event.message || 'The Visio edit worker failed.'));
			worker.onmessage = (event) => {
				try {
					const value = event.data;
					if (
						value?.ok === true &&
						value.bytes instanceof Uint8Array &&
						value.bytes.byteLength <= MAX_INPUT_BYTES &&
						value.document &&
						Array.isArray(value.changedParts) &&
						value.changedParts.length <= 1000 &&
						value.changedParts.every(
							(part: unknown) => typeof part === 'string' && part.length <= 4096,
						) &&
						Array.isArray(value.diagnostics) &&
						value.diagnostics.length <= 1000 &&
						value.diagnostics.every(
							(note: unknown) =>
								!!note &&
								typeof note === 'object' &&
								'code' in note &&
								typeof note.code === 'string' &&
								note.code.length <= 256 &&
								'message' in note &&
								typeof note.message === 'string' &&
								note.message.length <= 4096,
						)
					)
						finish(undefined, value as EditTransactionResult);
					else {
						const error = new Error(
							typeof value?.message === 'string' && value.message.length <= 4096
								? value.message
								: 'The Visio edit worker returned an invalid response.',
						);
						if (typeof value?.code === 'string' && value.code.length <= 256)
							Object.assign(error, { code: value.code });
						finish(error);
					}
				} catch (cause) {
					finish(cause instanceof Error ? cause : new Error(String(cause)));
				}
			};
			try {
				worker.postMessage({ bytes: copy.buffer, edits: commands }, [copy.buffer]);
			} catch (cause) {
				finish(cause instanceof Error ? cause : new Error(String(cause)));
			}
		});
	};
	edit.cancel = () => cancel?.();
	return edit;
}
