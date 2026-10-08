// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { beforeAll, describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import {
	createRuler,
	markerChange,
	markerPositions,
	updateRuler,
	type RulerGeometry,
} from './ruler';
import { rulerGeometry, syncRuler } from './ruler-sync';
import { schema } from './schema';
import { registerOfficeUi } from '../index';

beforeAll(() => registerOfficeUi());

const geometry = (over: Partial<RulerGeometry> = {}): RulerGeometry => ({
	pageWidth: 816,
	marginLeft: 96,
	marginRight: 96,
	indentLeft: 0,
	indentRight: 0,
	firstLine: 0,
	...over,
});

describe('ruler geometry', () => {
	it('places the markers from the margins and the paragraph indents', () => {
		expect(markerPositions(geometry())).toEqual({ firstLine: 96, left: 96, right: 720 });
		expect(markerPositions(geometry({ indentLeft: 48, indentRight: 24, firstLine: 24 }))).toEqual({
			firstLine: 168,
			left: 144,
			right: 696,
		});
		expect(markerPositions(geometry({ indentLeft: 48, firstLine: -48 })).firstLine).toBe(96);
	});
});

const marker = (ruler: HTMLElement, name: string) =>
	ruler.shadowRoot!.querySelector<HTMLElement>(`[data-marker="${name}"]`)!;

describe('ruler element', () => {
	it('is the shared ruler, sized to the page, with shaded margins and a zoom', () => {
		const ruler = createRuler();
		document.body.append(ruler);
		updateRuler(ruler, geometry({ indentLeft: 48 }), 1.5);
		expect(ruler.localName).toBe('office-ui-ruler');
		expect(ruler.classList.contains('dve-ruler')).toBe(true);
		expect(ruler.getAttribute('aria-label')).toBe('Ruler');
		expect(ruler.style.width).toBe('816px');
		expect(ruler.style.getPropertyValue('zoom')).toBe('1.5');
		const bands = ruler.shadowRoot!.querySelectorAll<HTMLElement>('.margin');
		expect([...bands].map((band) => band.style.width)).toEqual(['96px', '96px']);
		expect(marker(ruler, 'left').style.left).toBe('144px');
		ruler.remove();
	});
});

describe('ruler sync', () => {
	function editor(attrs: Record<string, unknown>) {
		const doc = schema.node('doc', { pageWidth: 816, marginLeft: 96, marginRight: 96 }, [
			schema.nodes.paragraph!.create({ id: 'p', ...attrs }, schema.text('text')),
		]);
		return new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
	}

	it("reads the caret paragraph's indents, hanging indent included", () => {
		const view = editor({ indentLeftTwips: 720, hangingTwips: 360, indentRightTwips: 150 });
		expect(rulerGeometry(view, createDocument())).toEqual({
			pageWidth: 816,
			marginLeft: 96,
			marginRight: 96,
			indentLeft: 48,
			indentRight: 10,
			firstLine: -24,
		});
	});

	it('updates a ruler in the frame and does nothing when there is none', () => {
		const frame = document.createElement('div');
		const ribbon = createRibbon();
		frame.append(ribbon);
		const view = editor({ indentLeftTwips: 720 });
		syncRuler(ribbon, view, createDocument());
		const ruler = createRuler();
		frame.append(ruler);
		syncRuler(ribbon, view, createDocument());
		expect(marker(ruler, 'left').style.left).toBe('144px');
	});

	it('has a View button that emits the toggle', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Ruler"]')!.click();
		expect(seen).toEqual([{ type: 'ruler' }]);
	});
});

describe('ruler dragging', () => {
	it('turns a dragged marker into indent changes snapped to sixteenths of an inch', () => {
		expect(markerChange('left', 96 + 48, geometry())).toEqual({ leftInches: 0.5 });
		expect(markerChange('left', 10, geometry())).toEqual({ leftInches: 0 });
		expect(markerChange('right', 720 - 97, geometry())).toEqual({ rightInches: 1.0 });
		expect(markerChange('right', 900, geometry())).toEqual({ rightInches: 0 });
		expect(markerChange('first-line', 96 + 24, geometry())).toEqual({
			special: 'firstLine',
			specialInches: 0.25,
		});
		expect(markerChange('first-line', 96 - 24, geometry({ indentLeft: 48 }))).toEqual({
			special: 'hanging',
			specialInches: 0.75,
		});
		expect(markerChange('first-line', 96, geometry())).toEqual({
			special: 'none',
			specialInches: 0,
		});
	});

	it('previews while dragging and reports one change on release', () => {
		const changes: unknown[] = [];
		const ruler = createRuler((change) => changes.push(change));
		document.body.append(ruler);
		updateRuler(ruler, geometry(), 1);
		const handle = marker(ruler, 'left');
		const fire = (type: string, clientX: number) =>
			handle.dispatchEvent(new MouseEvent(type, { clientX, bubbles: true, button: 0 }));
		fire('pointerdown', 96);
		fire('pointermove', 144);
		expect(marker(ruler, 'left').style.left).toBe('144px');
		expect(changes).toEqual([]);
		fire('pointerup', 144);
		expect(changes).toEqual([{ leftInches: 0.5 }]);
		ruler.remove();
	});
});
