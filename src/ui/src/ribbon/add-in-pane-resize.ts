const defaultWidth = 320;
const minimumWidth = 180;
const maximumWidth = 640;
const keyStep = 16;

export class AddInPaneWidth {
	value = defaultWidth;
	min = minimumWidth;
	max = maximumWidth;
	#drag: { width: number; start: number; direction: number } | undefined;

	constructor(
		private readonly host: HTMLElement,
		private readonly update: () => void,
	) {}

	layout() {
		const available = this.host.parentElement?.clientWidth;
		if (available) {
			this.max = Math.min(
				maximumWidth,
				available - Math.min(defaultWidth, Math.floor(available / 2)),
			);
			this.min = Math.min(minimumWidth, this.max);
		}
		this.resize(this.value);
	}

	private resize(width: number) {
		this.value = Math.min(this.max, Math.max(this.min, Math.round(width)));
		this.host.style.setProperty('--add-in-pane-width', `${this.value}px`);
		this.update();
	}

	private direction() {
		return this.host.ownerDocument.defaultView?.getComputedStyle(this.host).direction === 'rtl'
			? 1
			: -1;
	}

	key = (event: KeyboardEvent) => {
		const direction = this.direction();
		const value =
			event.key === 'Home'
				? this.min
				: event.key === 'End'
					? this.max
					: event.key === 'ArrowLeft'
						? this.value - keyStep * direction
						: event.key === 'ArrowRight'
							? this.value + keyStep * direction
							: undefined;
		if (value === undefined) return;
		event.preventDefault();
		event.stopPropagation();
		this.resize(value);
	};

	press = (event: PointerEvent) => {
		if (event.button !== 0 || !(event.currentTarget instanceof HTMLElement)) return;
		event.currentTarget.focus();
		event.currentTarget.setPointerCapture(event.pointerId);
		this.#drag = { width: this.value, start: event.clientX, direction: this.direction() };
	};

	move = (event: PointerEvent) => {
		const drag = this.#drag;
		if (drag) this.resize(drag.width + (event.clientX - drag.start) * drag.direction);
	};

	release = () => {
		this.#drag = undefined;
	};
}
