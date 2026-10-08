import { afterEach, expect, it } from 'vitest';
import { registerControls } from '../controls';
import type { OfficeUiFindBar } from './find-bar';

registerControls();
afterEach(() => document.body.replaceChildren());
async function setup(replaceMode = true) {
	const bar = document.createElement('office-ui-find-bar') as OfficeUiFindBar;
	document.body.append(bar);
	bar.replaceMode = replaceMode;
	bar.scopeOptions = [
		{ value: 'selected', label: 'Selection' },
		{ value: 'document', label: 'Document' },
	];
	bar.scope = 'selected';
	bar.matchCase = true;
	bar.matchCaseDisabled = true;
	bar.show();
	await bar.updateComplete;
	return { bar, root: bar.shadowRoot! };
}

it('adds opt-in inert replacement inputs and typed requests without changing ordinary Find', async () => {
	const ordinary = await setup(false);
	expect(ordinary.root.querySelector('textarea')).toBeNull();
	expect(ordinary.root.querySelector('[data-action="replace"]')).toBeNull();
	const { bar, root } = await setup();
	const events: unknown[] = [];
	bar.addEventListener('office-find-options', (event) =>
		events.push((event as CustomEvent).detail),
	);
	bar.addEventListener('office-find-replace', (event) =>
		events.push((event as CustomEvent).detail),
	);
	const input = root.querySelector('textarea')!;
	input.value = ' <&> Ω\n\n';
	input.dispatchEvent(new Event('input'));
	expect(bar.replacement).toBe(' <&> Ω\n\n');
	root.querySelector<HTMLButtonElement>('[data-action="replace"]')!.click();
	root.querySelector<HTMLButtonElement>('[data-action="replace-all"]')!.click();
	expect(events).toEqual([
		{ replacement: ' <&> Ω\n\n', scope: 'selected', matchCase: true },
		{ mode: 'current' },
		{ mode: 'all' },
	]);
	expect(root.querySelector('script')).toBeNull();
});

it('keeps MatchCase checked and disabled, honors action disabled state and escapes from replacement input', async () => {
	const { bar, root } = await setup();
	const checkbox = root.querySelector<HTMLElement & { checked: boolean; disabled: boolean }>(
		'office-ui-checkbox',
	)!;
	expect(checkbox.checked).toBe(true);
	expect(checkbox.disabled).toBe(true);
	let requests = 0;
	bar.addEventListener('office-find-replace', () => ++requests);
	bar.replaceDisabled = true;
	bar.replaceAllDisabled = true;
	await bar.updateComplete;
	root.querySelector<HTMLButtonElement>('[data-action="replace"]')!.click();
	root.querySelector<HTMLButtonElement>('[data-action="replace-all"]')!.click();
	expect(requests).toBe(0);
	const replacement = root.querySelector('textarea')!;
	bar.value = 'query';
	bar.replacement = 'keep';
	replacement.dispatchEvent(
		new KeyboardEvent('keydown', {
			key: 'Escape',
			isComposing: true,
			bubbles: true,
			composed: true,
		}),
	);
	expect(bar.value).toBe('query');
	replacement.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }),
	);
	expect(bar.value).toBe('');
	expect(bar.open).toBe(true);
	replacement.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, composed: true }),
	);
	expect(bar.open).toBe(false);
	expect(bar.replacement).toBe('keep');
});

it('renders errors as text and keeps query navigation separate from replacement textarea Enter', async () => {
	const { bar, root } = await setup();
	let steps = 0;
	bar.addEventListener('office-find-step', () => ++steps);
	root
		.querySelector('textarea')!
		.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(steps).toBe(0);
	root
		.querySelector('input')!
		.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(steps).toBe(1);
	bar.error = '<script>bad</script>';
	await bar.updateComplete;
	expect(root.querySelector('[role="alert"]')!.textContent).toBe('<script>bad</script>');
	expect(root.querySelector('script')).toBeNull();
});
