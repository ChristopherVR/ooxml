/** Core error codes remain visible for a refused transaction; document content is plain text. */
export function editErrorMessage(error: unknown): string {
	const message = error instanceof Error ? error.message : String(error);
	const code =
		error && typeof error === 'object' && 'code' in error && typeof error.code === 'string'
			? error.code.slice(0, 256)
			: undefined;
	return code ? `${code}: ${message}` : message;
}

/** Superseded worker/load operations are cancellation, not a refused document edit. */
export function isEditCancellation(error: unknown): boolean {
	return !!error && typeof error === 'object' && 'name' in error && error.name === 'AbortError';
}
