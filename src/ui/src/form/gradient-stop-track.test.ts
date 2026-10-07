// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
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

it('previews crossing and clamping, commits once, and cancels interrupted gestures', () => {
	const track = createGradientStopTrack(document);
	document.body.append(track.element);
	const stops = [
		{ position: 20, color: '#FF0000' },
		{ position: 80, color: '#0000FF' },
	];
	const commits: Array<[number, number]> = [];
	const previews: Array<number | undefined> = [];
	let disabled = false;
	const update = () =>
		track.update({
			stops,
			selected: 0,
			disabled,
			label: 'Stops',
			stopLabel: (index) => `Stop ${index}`,
			onSelect: update,
			onMove: (index, value) => {
				commits.push([index, value]);
				stops[index]!.position = value;
				update();
			},
			onPreview: (_index, value) => {
				previews.push(value);
			},
		});
	update();
	vi.spyOn(
		track.element.querySelector<HTMLElement>('.office-gradient-stop-paint')!,
		'getBoundingClientRect',
	).mockReturnValue({ width: 200 } as DOMRect);
	const pointer = (type: string, x: number, target: Element = track.element, id = 7) => {
		const event = new MouseEvent(type, { clientX: x, button: 0, bubbles: true, cancelable: true });
		Object.defineProperty(event, 'pointerId', { value: id });
		target.dispatchEvent(event);
	};
	const down = () => pointer('pointerdown', 140, track.element.querySelector('button')!);
	down();
	pointer('pointermove', 270);
	expect(stops[0]!.position).toBe(20);
	expect(previews.at(-1)).toBe(85);
	pointer('pointerup', 270);
	expect(commits).toEqual([[0, 85]]);
	expect(stops[1]!.position).toBe(80);
	down();
	pointer('pointermove', 1000);
	expect(previews.at(-1)).toBe(100);
	pointer('pointermove', -1000);
	expect(previews.at(-1)).toBe(0);
	pointer('pointercancel', -1000, track.element, 99);
	expect(previews.at(-1)).toBe(0);
	track.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
	expect(previews.at(-1)).toBeUndefined();
	pointer('pointerup', -1000);
	expect(commits).toHaveLength(1);
	down();
	pointer('pointermove', 0);
	disabled = true;
	update();
	pointer('pointerup', 0);
	expect(commits).toHaveLength(1);
	disabled = false;
	update();
	down();
	pointer('pointermove', 0);
	pointer('pointercancel', 0);
	pointer('pointerup', 0);
	expect(commits).toHaveLength(1);
	down();
	pointer('pointermove', 0);
	track.element.remove();
	pointer('pointerup', 0);
	expect(commits).toHaveLength(1);
});
