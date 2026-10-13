export interface PrintDragScroll {
	update(x: number, y: number): void;
	stop(): void;
	destroy(): void;
}
/** Scroll only the document viewport; each tick extends the original editor selection. */
export function createPrintDragScroll(
	viewport: HTMLElement,
	selectAt: (x: number, y: number) => void,
): PrintDragScroll {
	let frame: number | undefined,
		point: { x: number; y: number } | undefined,
		lastTime = 0,
		destroyed = false;
	function bounds() {
		const rect = viewport.getBoundingClientRect();
		return {
			left: rect.left,
			right: rect.right,
			top: Math.max(0, rect.top),
			bottom: Math.min(window.innerHeight, rect.bottom),
		};
	}
	function velocity(y: number, top: number, bottom: number): number {
		const band = Math.min(32, (bottom - top) / 3);
		if (band <= 0) return 0;
		if (y < top + band) return -Math.min(24, (24 * (top + band - y)) / band);
		if (y > bottom - band) return Math.min(24, (24 * (y - bottom + band)) / band);
		return 0;
	}
	function tick(time: number) {
		frame = undefined;
		if (!point || destroyed) return;
		const box = bounds(),
			speed = velocity(point.y, box.top, box.bottom);
		if (!speed) return;
		const elapsed = lastTime ? Math.min(64, time - lastTime) : 16;
		lastTime = time;
		const old = viewport.scrollTop,
			max = Math.max(0, viewport.scrollHeight - viewport.clientHeight);
		viewport.scrollTop = Math.max(0, Math.min(max, old + (speed * elapsed) / 16));
		if (viewport.scrollTop === old) return;
		selectAt(
			Math.max(box.left + 1, Math.min(box.right - 1, point.x)),
			Math.max(box.top + 1, Math.min(box.bottom - 1, point.y)),
		);
		frame = requestAnimationFrame(tick);
	}
	return {
		update(x, y) {
			if (destroyed) return;
			point = { x, y };
			const box = bounds();
			if (frame === undefined && velocity(y, box.top, box.bottom)) {
				lastTime = 0;
				frame = requestAnimationFrame(tick);
			}
		},
		stop() {
			point = undefined;
			lastTime = 0;
			if (frame !== undefined) cancelAnimationFrame(frame);
			frame = undefined;
		},
		destroy() {
			this.stop();
			destroyed = true;
		},
	};
}
