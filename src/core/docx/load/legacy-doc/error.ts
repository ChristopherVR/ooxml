/** Raised when a Word 97-2003 .doc cannot be opened, or an edit cannot be written back to it. */
export class LegacyDocError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'LegacyDocError';
	}
}
