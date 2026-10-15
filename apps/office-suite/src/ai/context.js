/**
 * What a request may read: the attached files, saved first so the assistant sees what is on
 * screen, as bounded text views, text files attached as notes, and optionally the open Teams
 * channel's recent messages.
 */
import { inspect } from '../ai-operations.js';

export async function gatherContext({
	sessions,
	ids,
	notes = [],
	teams: client,
	maxChars = 150000,
}) {
	const docs = [],
		files = [];
	const list = [...ids];
	// Each file gets an equal share of what is left after the notes; a file that needs less
	// passes the rest on to the next.
	let left = maxChars - JSON.stringify(notes).length - 2000;
	if (left < 1000 * list.length && list.length)
		throw new Error('The attached notes leave no room for the files. Attach fewer files.');
	for (const [i, id] of list.entries()) {
		const doc = await sessions.save(id);
		if (!doc) continue;
		const { name, kind, revision, parent } = doc;
		docs.push({ id, name, kind, revision, parent });
		const content = await inspect(doc, Math.floor(left / (list.length - i)));
		left -= JSON.stringify(content).length;
		files.push({ id, name, kind, revision, content });
	}
	const teams = client
		? client
				.getState()
				.messages.slice(-50)
				.map((m) => ({
					author: m.authorName,
					text: m.text,
					attachments: m.attachments.map((a) => ({ name: a.name, kind: a.kind })),
				}))
		: undefined;
	const context = { files, teams, ...(notes.length ? { notes } : {}) };
	const source = JSON.stringify(context);
	if (source.length > maxChars)
		throw new Error('The attached files are too large. Attach fewer files.');
	return { docs, context: files.length || teams || notes.length ? context : null, source };
}
