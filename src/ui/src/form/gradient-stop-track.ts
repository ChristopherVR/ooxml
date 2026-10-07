import { sortGradientStops } from 'ooxml-core/chart';

export interface GradientStopTrackOptions {
	stops: readonly { position: number; color: string }[];
	selected: number;
	disabled: boolean;
	label: string;
	stopLabel(index: number): string;
	onSelect(index: number): void;
}

/** Native buttons provide keyboard navigation and focus for the shared Office gradient strip. */
export function createGradientStopTrack(doc: Document) {
	const element = doc.createElement('div');
	element.className = 'office-gradient-stop-track';
	element.setAttribute('role', 'group');
	const style = doc.createElement('style');
	style.textContent = `.office-gradient-stop-track{position:relative;height:40px;margin:12px 7px}
.office-gradient-stop-paint{position:absolute;inset:3px 0 15px;border:1px solid var(--line,#aaa)}
.office-gradient-stop{position:absolute;top:0;width:14px;height:30px;padding:0;border:1px solid var(--line,#888);border-radius:2px;cursor:pointer}
.office-gradient-stop[aria-pressed=true]{outline:2px solid var(--accent,#217346);outline-offset:2px}
.office-gradient-stop:focus-visible{outline:2px solid var(--accent,#217346);outline-offset:3px}
.office-gradient-stop:disabled{opacity:.5;cursor:default}`;
	const paint = doc.createElement('div');
	paint.className = 'office-gradient-stop-paint';
	let current: GradientStopTrackOptions;
	const update = (options: GradientStopTrackOptions) => {
		current = options;
		element.setAttribute('aria-label', options.label);
		paint.style.background = `linear-gradient(90deg,${sortGradientStops(options.stops)
			.map((stop) => `${stop.color} ${stop.position}%`)
			.join(',')})`;
		element.replaceChildren(
			style,
			paint,
			...options.stops.map((stop, index) => {
				const button = doc.createElement('button');
				button.type = 'button';
				button.className = 'office-gradient-stop';
				button.disabled = options.disabled;
				button.setAttribute('aria-label', options.stopLabel(index));
				button.setAttribute('aria-pressed', String(index === options.selected));
				button.style.left = `calc(${stop.position}% - 7px)`;
				button.style.background = stop.color;
				button.addEventListener('click', () => {
					if (!current.disabled) current.onSelect(index);
				});
				button.addEventListener('keydown', (event) => {
					const next =
						event.key === 'ArrowRight'
							? index + 1
							: event.key === 'ArrowLeft'
								? index - 1
								: undefined;
					if (next === undefined || current.disabled) return;
					event.preventDefault();
					const selected = Math.max(0, Math.min(current.stops.length - 1, next));
					current.onSelect(selected);
					element.querySelectorAll<HTMLButtonElement>('button')[selected]?.focus();
				});
				return button;
			}),
		);
	};
	return { element, update };
}
