const CONTROL = 'pptx-ui-ribbon-command,pptx-ui-select,button,input[type=range],summary';
const NATIVE_KEYS = [
	' ',
	'Enter',
	'ArrowLeft',
	'ArrowRight',
	'ArrowUp',
	'ArrowDown',
	'Home',
	'End',
	'PageUp',
	'PageDown',
];

export function handleDrawKey(host: HTMLElement, event: KeyboardEvent): void {
	const target = event.target;
	if (
		event.defaultPrevented ||
		!(target instanceof HTMLElement) ||
		!NATIVE_KEYS.includes(event.key)
	)
		return;
	if (target.matches('input[type=range],pptx-ui-select[open]')) {
		event.stopPropagation();
		return;
	}
	if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
	const step = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
	if (!step && event.key !== 'Home' && event.key !== 'End') return;
	const controls = [...host.querySelectorAll<HTMLElement>(CONTROL)].filter(
		(control) =>
			!control.matches('[disabled],[aria-disabled=true]') && control.getClientRects().length > 0,
	);
	const index = controls.indexOf(target);
	if (index < 0) return;
	const next =
		event.key === 'Home'
			? 0
			: event.key === 'End'
				? controls.length - 1
				: (index + step + controls.length) % controls.length;
	event.preventDefault();
	controls[next]?.focus();
}
