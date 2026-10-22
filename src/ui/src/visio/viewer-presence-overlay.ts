import type { VisioPage } from 'ooxml-core/visio';
import type { RemotePresence } from 'ooxml-core/collab';
import { visioShapeFrame, type VisioPresence } from 'ooxml-core/visio/ui';
import { matrix } from './render-svg';

const NS = 'http://www.w3.org/2000/svg';
const COLOR = /^#[\da-f]{3,8}$/i;
/** Cursor arrow in screen pixels, tip at the origin. */
const ARROW = 'M0 0 L0 14 L3.8 10.6 L6.4 16.6 L8.6 15.6 L6 9.8 L11 9.8 Z';

export type PeerPresence = RemotePresence<VisioPresence>;

/** Pixels per drawing unit of `svg`, from its screen matrix (96 per inch at 100% as fallback). */
export function presenceScale(svg: SVGSVGElement, zoom: number): number {
	const screen = svg.getScreenCTM?.();
	const scale = screen ? Math.hypot(screen.a, screen.b) : 96 * zoom;
	return scale > 0 && Number.isFinite(scale) ? scale : 96;
}

/**
 * Draws the collaborators on `page` into one `<g data-presence-overlay>` of the page SVG, in
 * drawing units so zoom and pan move it with the drawing; pixel sizes are divided by `scale`.
 * Replaces any previous overlay and removes it when nobody is on the page.
 */
export function renderPresenceOverlay(
	svg: SVGSVGElement,
	page: VisioPage,
	peers: readonly PeerPresence[],
	scale: number,
): void {
	svg.querySelector(':scope > [data-presence-overlay]')?.remove();
	const here = peers.filter((peer) => peer.page === page.id);
	if (!here.length) return;
	const create = <K extends keyof SVGElementTagNameMap>(name: K) =>
		svg.ownerDocument.createElementNS(NS, name) as SVGElementTagNameMap[K];
	const overlay = create('g');
	overlay.dataset.presenceOverlay = '';
	overlay.setAttribute('pointer-events', 'none');
	overlay.setAttribute('aria-hidden', 'true');
	for (const peer of here) {
		const color = COLOR.test(peer.userColor) ? peer.userColor : '#2563eb';
		for (const id of peer.selection ?? []) {
			const frame = visioShapeFrame(page, id);
			if (!frame) continue;
			const outline = create('g');
			outline.dataset.presenceSelection = id;
			outline.dataset.presenceClient = String(peer.clientId);
			outline.setAttribute(
				'transform',
				`translate(0 ${page.height}) scale(1 -1) ${matrix(frame.transform)}`,
			);
			const rectangle = create('rect');
			rectangle.setAttribute('width', String(frame.width));
			rectangle.setAttribute('height', String(frame.height));
			rectangle.setAttribute('fill', 'none');
			rectangle.setAttribute('stroke', color);
			rectangle.setAttribute('stroke-width', '2');
			rectangle.setAttribute('stroke-dasharray', '5 3');
			rectangle.setAttribute('vector-effect', 'non-scaling-stroke');
			outline.append(rectangle);
			overlay.append(outline);
		}
	}
	for (const peer of here) {
		if (peer.x === undefined || peer.y === undefined) continue;
		const color = COLOR.test(peer.userColor) ? peer.userColor : '#2563eb';
		const cursor = create('g');
		cursor.dataset.presenceCursor = String(peer.clientId);
		cursor.setAttribute('transform', `translate(${peer.x} ${peer.y}) scale(${1 / scale})`);
		const arrow = create('path');
		arrow.setAttribute('d', ARROW);
		arrow.setAttribute('fill', color);
		arrow.setAttribute('stroke', '#fff');
		arrow.setAttribute('stroke-width', '1');
		const label = create('g');
		label.setAttribute('transform', 'translate(12 16)');
		const text = create('text');
		text.dataset.presenceName = '';
		text.textContent = peer.userName;
		text.setAttribute('x', '6');
		text.setAttribute('y', '12');
		const chip = create('rect');
		chip.setAttribute('width', String(Math.min(160, 12 + peer.userName.length * 6.5)));
		chip.setAttribute('height', '17');
		chip.setAttribute('rx', '3');
		chip.setAttribute('fill', color);
		label.append(chip, text);
		cursor.append(arrow, label);
		overlay.append(cursor);
	}
	svg.append(overlay);
}
