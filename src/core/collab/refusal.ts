// What a collaboration session refuses, in one vocabulary every product shares, so each can tell its
// user why in their own language (the messages are keys of `collabStrings` in `ooxml-core/i18n`).
// New code.

/**
 * Why an update was refused.
 * - `too-large`: an update from the room is over the byte or allocation budget of one update.
 * - `malformed`: an update from the room is not a well-formed Yjs update.
 * - `server-rejected`: the server refused a change this client sent.
 * - `server-full`: the shared document is at the server's size limit.
 */
export type UpdateRefusalCode = 'too-large' | 'malformed' | 'server-rejected' | 'server-full';

/** The English message of each refusal, also its translation key. */
export const UPDATE_REFUSAL_MESSAGES: Readonly<Record<UpdateRefusalCode, string>> = {
	'too-large':
		'A change from the shared session is too large to apply here, so sharing stopped. Your copy is unchanged.',
	malformed:
		'A change from the shared session could not be read, so sharing stopped. Your copy is unchanged.',
	'server-rejected':
		'The server refused your latest change, so sharing stopped. Your copy keeps the change.',
	'server-full':
		'The shared document has reached the size limit of the server, so sharing stopped. Your copy keeps your latest change.',
};

/**
 * A refused update. `message` is the sentence to show, English unless a product passes its
 * translation (see {@link UPDATE_REFUSAL_MESSAGES}); `detail` says what was refused, for logs.
 */
export class UpdateRefusedError extends Error {
	override readonly name = 'UpdateRefusedError';
	constructor(
		readonly code: UpdateRefusalCode,
		readonly detail: string,
		message: string = UPDATE_REFUSAL_MESSAGES[code],
	) {
		super(message);
	}
}

/** A value a collaborator wrote that this client left out, because its field does not allow it. */
export interface LeftOutValue {
	/** The part of the shared document it was written to, such as `links`. */
	readonly part: string;
	/** The field, or `entry` when the whole item was left out. */
	readonly field: string;
}

/** The English message shown when values a collaborator sent were left out; its translation key. */
export const LEFT_OUT_MESSAGE =
	'Content a collaborator sent was left out because it is not allowed here, such as an unsafe link or file type.';

/**
 * Whether `error` is an {@link UpdateRefusedError}, also one made by another copy of this package
 * (a host's own session), which `instanceof` would miss.
 */
export function isUpdateRefusal(error: unknown): error is UpdateRefusedError {
	return (
		error instanceof UpdateRefusedError ||
		(error instanceof Error &&
			error.name === 'UpdateRefusedError' &&
			Object.hasOwn(UPDATE_REFUSAL_MESSAGES, (error as { code?: unknown }).code as string))
	);
}
