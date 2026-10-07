// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createGradientStopTrack } from './gradient-stop-track';

it('keeps stop identities and keyboard focus, clamps navigation and suppresses disabled callbacks', () => {
	const track = createGradientStopTrack(document);
	document.body.append(track.element);
	let selected = 0;
	let disabled = false;
	const refresh = () =>
		track.update({
			stops: [
				{ position: 80, color: '#FF0000' },
				{ position: 20, color: '#0000FF' },
			],
			selected,
			disabled,
			label: 'Gradient stops',
			stopLabel: (index) => `Stop ${index + 1}`,
			onSelect: (index) => {
				selected = index;
				refresh();
			},
		});
	refresh();
	const buttons = () => track.element.querySelectorAll<HTMLButtonElement>('button');
	buttons()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }));
	expect(selected).toBe(1);
	expect(document.activeElement).toBe(buttons()[1]);
	buttons()[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }));
	expect(selected).toBe(1);
	disabled = true;
	refresh();
	buttons()[0]!.dispatchEvent(new Event('click'));
	expect(selected).toBe(1);
	expect(
		track.element.querySelector<HTMLElement>('.office-gradient-stop-paint')!.style.background,
	).toContain('rgb(0, 0, 255) 20%, rgb(255, 0, 0) 80%');
	track.element.remove();
});
