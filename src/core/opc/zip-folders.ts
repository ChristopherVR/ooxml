import type JSZip from 'jszip';

/** The entry names of a package, taken when it is opened, before anything is written to it. */
export const entryNames = (zip: JSZip): ReadonlySet<string> => new Set(Object.keys(zip.files));

/**
 * Removes the folder entries JSZip adds for the parent folders of each file it writes (its
 * `createFolders` default) unless the package had them when it was opened (`source`, from
 * {@link entryNames}). Office writes no folder entries, so a saved package keeps the shape of its
 * source instead of gaining a `word/` or `xl/` entry on every save.
 */
export function dropAddedFolders(zip: JSZip, source: ReadonlySet<string>): void {
	for (const [name, entry] of Object.entries(zip.files))
		if (entry.dir && !source.has(name)) delete zip.files[name];
}
