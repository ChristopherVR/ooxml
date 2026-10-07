/** Place native top-layer content beside its trigger, inside the owning viewport. */
export function toggleAnchoredPopover(panel: HTMLElement, trigger: HTMLElement): boolean {
	if (panel.matches(':popover-open')) {
		panel.hidePopover();
		return false;
	}
	panel.showPopover();
	const button = trigger.getBoundingClientRect();
	const bounds = panel.getBoundingClientRect();
	const view = panel.ownerDocument.defaultView!;
	panel.style.inset = 'auto';
	panel.style.margin = '0';
	panel.style.left = `${Math.max(0, Math.min(button.right - bounds.width, view.innerWidth - bounds.width))}px`;
	panel.style.top = `${Math.max(0, Math.min(button.bottom, view.innerHeight - bounds.height))}px`;
	return true;
}
