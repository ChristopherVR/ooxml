// Small browser helpers of the Calendar app.

/** Offer text as a file download. Returns false when there is no DOM to do it in. */
export function downloadText(name: string, text: string, type = 'text/calendar'): boolean {
	const doc = globalThis.document;
	if (!doc || typeof URL.createObjectURL !== 'function') return false;
	const url = URL.createObjectURL(new Blob([text], { type: `${type};charset=utf-8` }));
	const link = doc.createElement('a');
	link.href = url;
	link.download = name;
	doc.body.append(link);
	link.click();
	link.remove();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
	return true;
}
