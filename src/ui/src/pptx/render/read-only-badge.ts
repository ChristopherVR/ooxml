/**
 * The "Read-only" badge in the quick-access row, decided once for all five bindings.
 *
 * It shows whenever the deck cannot be edited (host `canEdit={false}`, a cloud file the write
 * policy leaves read-only, collaboration read-only), except in Protected View with an Enable
 * Editing action, where that action takes the badge's place. Bindings map the descriptor onto
 * their template; they keep no copy of the rule.
 */
export const READ_ONLY_BADGE_ATTR = 'data-pptx-read-only-badge';

export interface ReadOnlyBadgeInput {
	/** Whether editing is active for the open deck. */
	editable: boolean;
	/** Protected View with an Enable Editing action: that button replaces the badge. */
	enableEditingAvailable?: boolean;
}

export interface ReadOnlyBadgeView {
	visible: boolean;
	textKey: 'pptx.toolbar.readOnly';
	/** The badge is a status message, announced politely. */
	role: 'status';
}

export function resolveReadOnlyBadge(input: ReadOnlyBadgeInput): ReadOnlyBadgeView {
	return {
		visible: !input.editable && !input.enableEditingAvailable,
		textKey: 'pptx.toolbar.readOnly',
		role: 'status',
	};
}
