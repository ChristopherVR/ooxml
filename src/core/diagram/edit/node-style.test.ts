import { describe, it, expect } from 'vitest';

import type { DiagramNode } from '../model/node';
import { setDiagramNodeStyle, type DiagramStyledShape } from './node-style';
import type { DiagramEditableData } from './types';

function data(): DiagramEditableData<DiagramNode, DiagramStyledShape & { x: number }> {
	return {
		nodes: [
			{ id: '1', text: 'One' },
			{ id: '2', text: 'Two', style: { fillColor: '#000000', bold: true } },
		],
		drawingShapes: [{ id: 's', x: 0 }],
	};
}

describe('setDiagramNodeStyle', () => {
	it('sets a style on a node that had none', () => {
		const next = setDiagramNodeStyle(data(), '1', { fillColor: '#FF0000', fontColor: '#FFFFFF' });
		expect(next.nodes[0]!.style).toStrictEqual({ fillColor: '#FF0000', fontColor: '#FFFFFF' });
	});

	it('shallow-merges a partial onto an existing style', () => {
		const next = setDiagramNodeStyle(data(), '2', { italic: true, fillColor: '#00FF00' });
		expect(next.nodes[1]!.style).toStrictEqual({ fillColor: '#00FF00', bold: true, italic: true });
	});

	it('removes a key when the partial sets it to undefined', () => {
		const next = setDiagramNodeStyle(data(), '2', { bold: undefined });
		expect(next.nodes[1]!.style).toStrictEqual({ fillColor: '#000000' });
	});

	it('drops the style field entirely when no keys remain', () => {
		const next = setDiagramNodeStyle(data(), '2', { fillColor: undefined, bold: undefined });
		expect(next.nodes[1]!.style).toBeUndefined();
		expect(next.nodes[1]).not.toHaveProperty('style');
	});

	it('marks drawingShapes as stale when shape cannot be matched', () => {
		const next = setDiagramNodeStyle(data(), '1', { fillColor: '#FF0000' });
		expect(next.drawingShapes).toStrictEqual([]);
	});

	it('patches fillColor in-place when shapes can be matched positionally', () => {
		const input: DiagramEditableData<DiagramNode, DiagramStyledShape & { x: number }> = {
			nodes: [
				{ id: '1', text: 'One' },
				{ id: '2', text: 'Two', style: { fillColor: '#000000', bold: true } },
			],
			drawingShapes: [
				{ id: 's1', x: 0 },
				{ id: 's2', x: 20 },
			],
		};
		const next = setDiagramNodeStyle(input, '1', { fillColor: '#FF0000' });
		expect(next.drawingShapes).toHaveLength(2);
		expect(next.drawingShapes![0]!.fillColor).toBe('#FF0000');
	});

	it('preserves undefined drawingShapes for freshly inserted SmartArt', () => {
		const input: DiagramEditableData<DiagramNode, DiagramStyledShape & { x: number }> = {
			nodes: [{ id: '1', text: 'One' }],
		};
		const next = setDiagramNodeStyle(input, '1', { fillColor: '#FF0000' });
		expect(next.drawingShapes).toBeUndefined();
	});

	it('is immutable: the input is not mutated', () => {
		const input = data();
		const snapshot = JSON.parse(JSON.stringify(input));
		setDiagramNodeStyle(input, '1', { fillColor: '#FF0000' });
		expect(input).toStrictEqual(snapshot);
	});

	it('returns the same reference when the node id is not found', () => {
		const input = data();
		expect(setDiagramNodeStyle(input, 'missing', { fillColor: '#FF0000' })).toBe(input);
	});
});
