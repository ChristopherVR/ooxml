// Insert > Shapes: Word's gallery of shapes by family. Choosing one arms the draw tool.
import { SHAPE_PALETTE, newShape, shapeSvg } from 'ooxml-core/docx';
import { localeOf, translate } from './localization';
import { mountPopover } from './ribbon-popover';

/** A thumbnail of the shape in Word's default style, sized to the gallery cell. */
function thumbnail(preset: string): HTMLElement | undefined {
	const svg = shapeSvg(newShape(preset), 22, 22);
	if (!svg) return undefined;
	const parsed = new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
	return parsed.localName === 'svg'
		? (document.importNode(parsed, true) as unknown as HTMLElement)
		: undefined;
}

/**
 * Opens the shape gallery under `anchor`. `onPick` receives the preset; the caller then lets the
 * user draw it on the page. Arrow keys move through the cells and Enter chooses.
 */
export function openShapePicker(
	anchor: HTMLElement,
	onPick: (preset: string) => void,
	onClose: () => void,
): void {
	const locale = localeOf(anchor);
	const say = (text: string) => translate(locale, text as never);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover shape-picker';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', say('Shapes'));
	const cells: HTMLButtonElement[] = [];
	for (const group of SHAPE_PALETTE) {
		const section = document.createElement('section');
		const heading = document.createElement('h4');
		heading.textContent = say(group.name);
		const grid = document.createElement('div');
		grid.className = 'shape-picker-grid';
		for (const item of group.shapes) {
			const cell = document.createElement('button');
			cell.type = 'button';
			cell.dataset.preset = item.preset;
			cell.title = say(item.name);
			cell.setAttribute('aria-label', say(item.name));
			const picture = thumbnail(item.preset);
			if (picture) cell.append(picture);
			cell.addEventListener('mousedown', (event) => event.preventDefault());
			cell.addEventListener('click', () => {
				onPick(item.preset);
				onClose();
			});
			cell.addEventListener('keydown', (event) => {
				const at = cells.indexOf(cell);
				const step = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: 8, ArrowUp: -8 }[event.key];
				if (step === undefined) return;
				event.preventDefault();
				cells[Math.min(cells.length - 1, Math.max(0, at + step))]?.focus();
			});
			cells.push(cell);
			grid.append(cell);
		}
		section.append(heading, grid);
		pop.append(section);
	}
	if (mountPopover(anchor, pop, anchor)) cells[0]?.focus({ preventScroll: true });
}
