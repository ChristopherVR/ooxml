import type { DocumentSaveProgress } from 'ooxml-core/embedding';

/** The English texts of a save in progress; `{percent}` is replaced with a whole number. */
export const SAVE_PROGRESS_TEXT = {
	uploading: 'Saving… {percent}%',
	checking: 'Checking for changes…',
	downloading: 'Getting the latest version… {percent}%',
	downloadingUnknown: 'Getting the latest version…',
	merging: 'Combining edits…',
} as const;

const percent = (loaded: number, total: number) =>
	String(Math.max(0, Math.min(100, Math.round((loaded / total) * 100))));

/**
 * The status text for one step of a save (`DocumentSaveMeta.onProgress`), the same in every
 * product. `t` translates the English text before `{percent}` is filled in, so a dictionary keeps
 * the placeholder where its language wants it.
 */
export function saveProgressLabel(
	progress: DocumentSaveProgress,
	t: (text: string) => string = (text) => text,
): string {
	switch (progress.phase) {
		case 'uploading':
			return t(SAVE_PROGRESS_TEXT.uploading).replace(
				'{percent}',
				progress.total > 0 ? percent(progress.loaded, progress.total) : '0',
			);
		case 'checking':
			return t(SAVE_PROGRESS_TEXT.checking);
		case 'downloading':
			return progress.total
				? t(SAVE_PROGRESS_TEXT.downloading).replace(
						'{percent}',
						percent(progress.loaded, progress.total),
					)
				: t(SAVE_PROGRESS_TEXT.downloadingUnknown);
		case 'merging':
			return t(SAVE_PROGRESS_TEXT.merging);
	}
}

/**
 * Whether a new step is worth announcing to assistive technology: a change of phase, or 10 more
 * percent of the same transfer. Visible text can update on every step; a live region should not.
 */
export function saveProgressStepChanged(
	previous: DocumentSaveProgress | undefined,
	next: DocumentSaveProgress,
): boolean {
	if (!previous || previous.phase !== next.phase) return true;
	if (
		(next.phase === 'uploading' || next.phase === 'downloading') &&
		(previous.phase === 'uploading' || previous.phase === 'downloading') &&
		next.total &&
		previous.total
	)
		return (
			Math.floor((next.loaded / next.total) * 10) !==
			Math.floor((previous.loaded / previous.total) * 10)
		);
	return false;
}
