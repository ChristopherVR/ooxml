/** A custom-show action target to replace, or null to remove its jump. */
export type CustomShowActionRewrite = (id: string) => string | null;

/** Reads the id of a custom-show action; malformed or unrelated URLs are not actions. */
export function customShowActionId(action: string | undefined): string | undefined {
	if (!action || !/^ppaction:\/\/customshow(?:\?|$)/i.test(action)) return undefined;
	const match = action.match(/[?&]id=([^&]*)/i);
	if (!match) return undefined;
	try {
		return decodeURIComponent(match[1] as string);
	} catch {
		return undefined;
	}
}

/** Changes only the id query value, preserving return=true and any other action options. */
export function rewriteCustomShowAction(
	action: string | undefined,
	rewrite: CustomShowActionRewrite,
) {
	const id = customShowActionId(action);
	if (id === undefined) return { changed: false, action };
	const next = rewrite(id);
	if (next === id) return { changed: false, action };
	return {
		changed: true,
		action:
			next === null
				? undefined
				: (action as string).replace(/([?&]id=)[^&]*/i, '$1' + encodeURIComponent(next)),
	};
}
