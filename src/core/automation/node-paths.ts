// How the file tools split and judge the text of a path. Internal to `automation/node`.

/**
 * A Windows path segment that does not name the file it spells: an alternate data stream or drive
 * (`:`), characters Windows rejects or expands, a trailing dot or space Windows drops, and the
 * reserved device names, with or without an extension.
 */
const WINDOWS_UNSAFE_SEGMENT =
	/[\u0000-\u001f:<>"|?*]|[. ]$|^(?:con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³]|conin\$|conout\$)(?:[. ].*)?$/iu;

/** Whether `segment` is refused on `platform`; POSIX file names may hold anything but NUL. */
export const isUnsafePathSegment = (segment: string, platform = process.platform): boolean =>
	platform === 'win32' ? WINDOWS_UNSAFE_SEGMENT.test(segment) : segment.includes('\0');

/** The segments of `path`; a backslash separates them only on Windows, elsewhere it is a name character. */
export const pathSegments = (path: string, platform = process.platform): string[] =>
	path.split(platform === 'win32' ? /[\\/]/ : /\//);
