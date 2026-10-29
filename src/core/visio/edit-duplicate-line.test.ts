import { describe, expect, it } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import {
	assertConnectorCopyApart,
	assertGlueDroppable,
	finishLineCopy,
	isLineSheet,
	planLineCopy,
} from './edit-duplicate-line';

const NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const cell = (n: string, v: number, f?: string) =>
	`<Cell N="${n}" V="${v}"${f ? ` F="${f.replace(/"/g, '&quot;')}"` : ''}/>`;
const glue = (id: string, index = 1) =>
	`PAR(PNT(Sheet.${id}!Connections.X${index},Sheet.${id}!Connections.Y${index}))`;
const line = (id: string) =>
	`<Shape ID="${id}" Type="Shape">${cell('PinX', 4, '(BeginX+EndX)/2')}${cell('PinY', 4.5, '(BeginY+EndY)/2')}` +
	`${cell('Width', Math.hypot(4, 2), 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)')}${cell('Height', 0)}` +
	`${cell('Angle', Math.atan2(-2, 4), 'ATAN2(EndY-BeginY,EndX-BeginX)')}${cell('LocPinX', Math.hypot(4, 2) / 2, 'Width*0.5')}` +
	`${cell('LocPinY', 0, 'Height*0.5')}${cell('BeginX', 2, glue('1'))}${cell('BeginY', 5.5, glue('1'))}` +
	`${cell('EndX', 6, glue('2'))}${cell('EndY', 3.5, glue('2'))}` +
	`<Section N="Geometry" IX="0"><Row T="MoveTo" IX="1">${cell('X', 0, 'Width*0')}${cell('Y', 0)}</Row>` +
	`<Row T="LineTo" IX="2">${cell('X', Math.hypot(4, 2), 'Width*1')}${cell('Y', 0)}</Row></Section></Shape>`;
const root = (shapes: string, connects = '') =>
	new DOMParser().parseFromString(
		`<PageContents xmlns="${NS}"><Shapes>${shapes}</Shapes>${connects}</PageContents>`,
		'text/xml',
	).documentElement as unknown as Element;
const shape = (page: Element) => page.getElementsByTagName('Shape')[0]! as unknown as Element;
const value = (node: Element, name: string) =>
	Array.from(node.getElementsByTagName('Cell') as unknown as Element[])
		.find((c) => c.getAttribute('N') === name)!
		.getAttribute('V');
const formula = (node: Element, name: string) =>
	Array.from(node.getElementsByTagName('Cell') as unknown as Element[])
		.find((c) => c.getAttribute('N') === name)!
		.getAttribute('F');

describe('copying a glued line', () => {
	it('recognizes a line and plans its pin', () => {
		const page = root(line('3'));
		expect(isLineSheet(shape(page))).toBe(true);
		expect(planLineCopy(shape(page))).toEqual({ x: 4, y: 4.5 });
	});
	it('drops the glue formulas and moves the ends and midpoint', () => {
		const copy = shape(root(line('4')));
		finishLineCopy(copy, 0.5, -1);
		expect(formula(copy, 'BeginX')).toBeNull();
		expect(value(copy, 'BeginX')).toBe('2.5');
		expect(value(copy, 'EndY')).toBe('2.5');
		expect(value(copy, 'PinX')).toBe('4.5');
		expect(formula(copy, 'PinX')).toBe('(BeginX+EndX)/2');
	});
	it('refuses a connector copied together with a shape it is glued to', () => {
		const page = root(
			line('3'),
			'<Connects><Connect FromSheet="3" ToSheet="1"/><Connect FromSheet="3" ToSheet="2"/></Connects>',
		);
		expect(() => assertGlueDroppable(page, '3', new Set(['3']))).not.toThrow();
		expect(() => assertGlueDroppable(page, '3', new Set(['3', '1']))).toThrow(/glued to/);
	});
	it('refuses a stencil connector copy that meets another dynamic connector', () => {
		const walk = (id: string, y: number) =>
			`<Shape ID="${id}">${cell('BeginX', 2, '_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)')}${cell('BeginY', y)}${cell('EndX', 5)}${cell('EndY', y)}</Shape>`;
		const page = root(walk('4', 5) + walk('9', 8));
		const source = page.getElementsByTagName('Shape')[1]! as unknown as Element;
		expect(() => assertConnectorCopyApart(page, [source], 0, -1)).not.toThrow();
		expect(() => assertConnectorCopyApart(page, [source], 0, -2.9)).toThrow(/lays out again/);
	});
});
