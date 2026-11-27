// The Draw tab's ruler stencil on the shared ink canvas (Word, Excel): a straightedge laid over
// the surface that you drag, rotate with the mouse wheel or the keyboard, and draw along, as in
// Office. A stroke begun near its edge follows the edge in a straight line. Coordinates are the
// canvas's own at 100% zoom, as its samples are.
import { svg, type TemplateResult } from 'lit';
import { translateInkText } from 'ooxml-core/i18n';

export interface InkPoint {
	x: number;
	y: number;
}

/** Height of the ruler's body, and how near its edge a stroke must begin to follow it. */
export const RULER_WIDTH = 48;
export const RULER_SNAP = 16;
const MM = 96 / 25.4;
const MAX_LENGTH = 900;

/** Degrees in -180..180. */
const wrap = (degrees: number): number => {
	const turned = ((((degrees + 180) % 360) + 360) % 360) - 180;
	return Object.is(turned, -0) ? 0 : turned;
};

export class InkRuler {
	angle = 0;
	center: InkPoint | undefined;
	length = 600;

	/** The ends of the drawing edge. */
	edge(): [InkPoint, InkPoint] {
		const c = this.center ?? { x: 0, y: 0 };
		const radians = (this.angle * Math.PI) / 180;
		const half = this.length / 2;
		const dx = Math.cos(radians) * half;
		const dy = Math.sin(radians) * half;
		return [
			{ x: c.x - dx, y: c.y - dy },
			{ x: c.x + dx, y: c.y + dy },
		];
	}

	/** Where `point` falls on the edge line: `offset` 0..1 along it, `distance` across it. */
	#project(point: InkPoint): { point: InkPoint; offset: number; distance: number } {
		const [a, b] = this.edge();
		const dx = b.x - a.x;
		const dy = b.y - a.y;
		const squared = dx * dx + dy * dy || 1;
		const offset = ((point.x - a.x) * dx + (point.y - a.y) * dy) / squared;
		const on = { x: a.x + offset * dx, y: a.y + offset * dy };
		return { point: on, offset, distance: Math.hypot(point.x - on.x, point.y - on.y) };
	}

	/** Whether a stroke begun at `point` follows the edge. */
	near(point: InkPoint): boolean {
		const { offset, distance } = this.#project(point);
		return offset >= 0 && offset <= 1 && distance <= RULER_SNAP;
	}

	/** `point` moved onto the edge (kept within its ends). */
	snap<T extends InkPoint>(point: T): T {
		const [a, b] = this.edge();
		const { offset } = this.#project(point);
		const t = Math.max(0, Math.min(1, offset));
		return { ...point, x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
	}

	rotate(by: number): void {
		this.angle = wrap(Math.round((this.angle + by) * 10) / 10);
	}

	moveTo(point: InkPoint): void {
		this.center = { ...point };
	}

	/** Places a ruler shown for the first time in the middle of what is visible of `box`. */
	place(visible: { x: number; y: number; width: number; height: number }): void {
		if (this.center) return;
		this.center = { x: visible.x + visible.width / 2, y: visible.y + visible.height / 2 };
		this.length = Math.max(160, Math.min(MAX_LENGTH, visible.width * 0.8 || 600));
	}

	/**
	 * Arrow keys move the ruler (Ctrl by a pixel), Shift with an arrow rotates it a degree (with
	 * Ctrl fifteen). True when the key was the ruler's.
	 */
	key(event: KeyboardEvent): boolean {
		const delta = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[
			event.key
		];
		if (!delta || event.altKey || event.metaKey) return false;
		const [dx = 0, dy = 0] = delta;
		if (event.shiftKey) this.rotate((dx + dy) * (event.ctrlKey ? 15 : 1));
		else {
			const step = event.ctrlKey ? 1 : 10;
			const c = this.center ?? { x: 0, y: 0 };
			this.moveTo({ x: c.x + dx * step, y: c.y + dy * step });
		}
		return true;
	}

	/** The mouse wheel turns the ruler a degree a notch (Shift: fifteen). */
	wheel(event: WheelEvent): void {
		const notch = Math.sign(event.deltaY || event.deltaX);
		if (notch) this.rotate(notch * (event.shiftKey ? 15 : 1));
	}

	/** The ruler's accessible name, with its angle. */
	label(locale: string | undefined): string {
		return `${translateInkText('Ruler', locale)}, ${translateInkText('{angle} degrees', locale, {
			angle: Math.round(this.angle),
		})}`;
	}

	/** The ruler drawn at `scale`, its edge on top of its body (before rotation). */
	render(
		scale: number,
		locale: string | undefined,
		handlers: { keydown: (event: KeyboardEvent) => void; wheel: (event: WheelEvent) => void },
	): TemplateResult {
		const c = this.center ?? { x: 0, y: 0 };
		const half = this.length / 2;
		let ticks = '';
		const numbers: TemplateResult[] = [];
		for (let i = 0, x = -half; x <= half; i++, x = -half + i * MM) {
			const size = i % 10 === 0 ? 14 : i % 5 === 0 ? 9 : 5;
			ticks += `M${x.toFixed(2)} 0V${size}`;
			if (i % 10 === 0 && i > 0 && x < half - 8)
				numbers.push(svg`<text class="ruler-number" x=${x} y="26">${i / 10}</text>`);
		}
		return svg`<g
			class="ruler"
			part="ruler"
			tabindex="0"
			role="group"
			aria-label=${this.label(locale)}
			aria-description=${translateInkText(
				'Drag or use the arrow keys to move. Hold Shift and press an arrow key, or turn the mouse wheel, to rotate.',
				locale,
			)}
			transform=${`translate(${c.x * scale} ${c.y * scale}) rotate(${this.angle}) scale(${scale})`}
			@keydown=${handlers.keydown}
			@wheel=${handlers.wheel}
		>
			<rect class="ruler-body" x=${-half} y="0" width=${this.length} height=${RULER_WIDTH} rx="2"></rect>
			<path class="ruler-ticks" d=${ticks}></path>
			${numbers}
			<text class="ruler-angle" x="0" y=${RULER_WIDTH - 8}>${Math.round(this.angle)}°</text>
		</g>`;
	}
}
