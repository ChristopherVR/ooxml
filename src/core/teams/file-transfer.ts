export interface FileTransferProgress {
	phase: 'preparing' | 'uploading' | 'uploaded';
	fileName: string;
	completed: number;
	total: number;
}

/** Progress counts completed files, not transmitted or durably stored bytes. */
export interface FileOperationOptions {
	signal?: AbortSignal;
	onProgress?: (progress: FileTransferProgress) => void;
}

export function checkFileAbort(signal?: AbortSignal): void {
	if (signal?.aborted) throw new DOMException('File sharing canceled', 'AbortError');
}

/** Stop waiting for adapters that ignore cancellation, while handling their late rejections. */
export function withFileAbort<T>(task: () => Promise<T>, signal?: AbortSignal): Promise<T> {
	return new Promise((resolve, reject) => {
		const abort = () => {
			signal?.removeEventListener('abort', abort);
			reject(new DOMException('File sharing canceled', 'AbortError'));
		};
		if (signal?.aborted) return abort();
		signal?.addEventListener('abort', abort, { once: true });
		try {
			task().then(
				(value) => {
					signal?.removeEventListener('abort', abort);
					resolve(value);
				},
				(error: unknown) => {
					signal?.removeEventListener('abort', abort);
					reject(error);
				},
			);
		} catch (error) {
			signal?.removeEventListener('abort', abort);
			reject(error);
		}
	});
}
