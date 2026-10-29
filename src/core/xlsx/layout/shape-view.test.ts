import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/index';
import type { FormControlObject, ShapeNode, UnsupportedDrawingObject } from '../model';
import {
	applyControlAction,
	controlState,
	formControlSvg,
	radioGroup,
	shapeSvg,
	wrapShapeLine,
} from './index';

const load = async () =>
	loadXlsx(
		new Uint8Array(
			readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-shapes.xlsx')),
		),
	);

const shapeByName = async (name: string): Promise<{ node: ShapeNode; theme: never }> => {
	const wb = await load();
	const found = wb.sheets[0]!.drawings.find(
		(d): d is UnsupportedDrawingObject => d.kind === 'unsupported' && d.shape?.name === name,
	);
	return { node: found!.shape!, theme: wb.theme as never };
};

describe('shapeSvg', () => {
	it('draws a rectangle with its fill, outline and text', async () => {
		const { node, theme } = await shapeByName('Box');
		const svg = shapeSvg(node, 160, 80, { theme });
		expect(svg.toLowerCase()).toContain('fill="#ff8000"');
		expect(svg).toContain('stroke="#000000"');
		expect(svg).toContain('stroke-width="2.67"');
		expect(svg).toContain('>Hello</text>');
	});

	it('takes an unset fill and outline from the style references and the theme', async () => {
		const { node, theme } = await shapeByName('Oval');
		const svg = shapeSvg(node, 130, 80, { theme });
		expect(svg).toMatch(/fill="#[0-9a-f]{6}"/);
		expect(svg).toMatch(/stroke="#[0-9a-f]{6}"/);
		expect(svg).toContain('A');
	});

	it('draws a preset arrow from the shared geometry', async () => {
		const { node, theme } = await shapeByName('Arrow');
		const svg = shapeSvg(node, 130, 80, { theme });
		expect(svg).toMatch(/<path d="M[^"]+Z"/);
	});

	it('leaves a text box without fill from its style and keeps its text', async () => {
		const { node, theme } = await shapeByName('Note');
		const svg = shapeSvg(node, 180, 52, { theme });
		expect(svg).toContain('A text box');
	});

	it('draws a connector as an unfilled line and a group from its children', async () => {
		const line = await shapeByName('Line');
		expect(shapeSvg(line.node, 210, 50, { theme: line.theme })).toContain('fill="none"');
		const group = await shapeByName('Pair');
		const svg = shapeSvg(group.node, 147, 40, { theme: group.theme });
		expect(svg.match(/<path /g)?.length).toBe(2);
	});

	it('wraps long lines to the box width', () => {
		expect(wrapShapeLine('one two three four five six', 60, 14).length).toBeGreaterThan(1);
	});
});

describe('form controls', () => {
	const controls = async () =>
		(await load()).sheets[0]!.drawings.filter(
			(d): d is FormControlObject => d.kind === 'formControl',
		);

	it('draws each control type from its linked value', async () => {
		const list = await controls();
		const by = (type: string) => list.find((c) => c.controlType === type)!;
		const check = formControlSvg(by('checkbox'), controlState(by('checkbox'), true), 120, 20)!;
		expect(check).toContain('Enable');
		expect(check).toContain('data-control-part="toggle"');
		expect(check).toContain('<path d="M');
		const unchecked = formControlSvg(by('checkbox'), controlState(by('checkbox'), false), 120, 20)!;
		expect(unchecked).not.toContain('stroke-width="1.6"');
		const drop = formControlSvg(by('dropdown'), controlState(by('dropdown'), 2), 100, 20, {
			items: ['Red', 'Green', 'Blue'],
		})!;
		expect(drop).toContain('>Green</text>');
		const box = formControlSvg(by('list'), controlState(by('list'), 3), 100, 65, {
			items: ['Red', 'Green', 'Blue'],
		})!;
		expect(box).toContain('fill="#0078d7"');
		expect(box.match(/data-control-index/g)).toHaveLength(3);
		expect(formControlSvg(by('button'), controlState(by('button'), undefined), 90, 24)).toContain(
			'Press me',
		);
		expect(formControlSvg(by('spin'), controlState(by('spin'), 3), 20, 40)).toContain('step-up');
		expect(formControlSvg(by('scroll'), controlState(by('scroll'), 40), 100, 16)).toContain(
			'thumb',
		);
		expect(
			formControlSvg(by('groupBox'), controlState(by('groupBox'), undefined), 120, 80),
		).toContain('Group');
	});

	it('computes the value an interaction writes to the linked cell', async () => {
		const list = await controls();
		const check = list[0]!;
		expect(applyControlAction(check, controlState(check, true), { type: 'toggle' })).toBe(false);
		expect(applyControlAction(check, controlState(check, false), { type: 'toggle' })).toBe(true);
		const [, one, two] = list;
		const group = radioGroup(list, two!);
		expect(group.buttons).toEqual([one, two]);
		expect(group.index).toBe(1);
		expect(group.linkedCell).toBe('$C$2');
		expect(applyControlAction(two!, controlState(two!, 1, 1), { type: 'toggle' }, 1)).toBe(2);
		expect(controlState(one!, 1, 0).checked).toBe('checked');
		expect(controlState(two!, 1, 1).checked).toBe('unchecked');
		const drop = list.find((c) => c.controlType === 'dropdown')!;
		expect(applyControlAction(drop, controlState(drop, 1), { type: 'select', index: 3 })).toBe(3);
		const spin = list.find((c) => c.controlType === 'spin')!;
		expect(applyControlAction(spin, controlState(spin, 10), { type: 'step', delta: 1 })).toBe(10);
		expect(applyControlAction(spin, controlState(spin, 3), { type: 'step', delta: -1 })).toBe(2);
		const scroll = list.find((c) => c.controlType === 'scroll')!;
		expect(
			applyControlAction(scroll, controlState(scroll, 95), { type: 'page', direction: 1 }),
		).toBe(100);
		expect(applyControlAction(scroll, controlState(scroll, 40), { type: 'set', value: -5 })).toBe(
			0,
		);
		const button = list.find((c) => c.controlType === 'button')!;
		expect(
			applyControlAction(button, controlState(button, undefined), { type: 'toggle' }),
		).toBeUndefined();
	});
});
