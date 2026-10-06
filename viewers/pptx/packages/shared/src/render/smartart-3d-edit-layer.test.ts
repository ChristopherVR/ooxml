// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';

import {
	routeEditLayerPointerToNodes,
	smartArtNodeAtPoint,
	stripEditLayerMarkers,
} from './smartart-3d-edit-layer';

describe('stripEditLayerMarkers', () => {
	it('removes element markers but keeps node ids', () => {
		const root = document.createElement('div');
		root.innerHTML =
			'<div data-element-id="e1" data-testid="smartart-list" role="img" aria-label="Diagram">' +
			'<svg><g data-smartart-node-id="n1"><text>A</text></g></svg></div>';
		root.setAttribute('data-element-id', 'e1');
		stripEditLayerMarkers(root);
		expect(root.hasAttribute('data-element-id')).toBeFalsy();
		expect(root.querySelector('[data-element-id], [data-testid], [role], [aria-label]')).toBeNull();
		expect(root.querySelector('[data-smartart-node-id="n1"]')).not.toBeNull();
	});

	it('keeps the accessible name of the controls the layer shows', () => {
		const root = document.createElement('div');
		root.innerHTML =
			'<div role="group" aria-label="Fill Color"><button aria-label="Fill Color #fff"></button></div>' +
			'<textarea aria-label="Edit node text"></textarea>';
		stripEditLayerMarkers(root);
		expect(root.querySelector('button')?.getAttribute('aria-label')).toBe('Fill Color #fff');
		expect(root.querySelector('textarea')?.getAttribute('aria-label')).toBe('Edit node text');
	});
});

describe('smartArtNodeAtPoint', () => {
	function node(id: string, left: number, top: number, width: number, height: number): Element {
		const el = document.createElement('div');
		el.setAttribute('data-smartart-node-id', id);
		el.getBoundingClientRect = () =>
			({ left, top, width, height, right: left + width, bottom: top + height }) as DOMRect;
		return el;
	}

	it('finds the smallest node box under the point, whatever its pointer-events', () => {
		const root = document.createElement('div');
		root.append(
			node('outer', 0, 0, 200, 200),
			node('inner', 50, 50, 40, 40),
			node('far', 300, 0, 50, 50),
		);
		expect(smartArtNodeAtPoint(root, 60, 60)?.getAttribute('data-smartart-node-id')).toBe('inner');
		expect(smartArtNodeAtPoint(root, 10, 10)?.getAttribute('data-smartart-node-id')).toBe('outer');
		expect(smartArtNodeAtPoint(root, 250, 250)).toBeNull();
	});

	/** A node `<g>` whose one shape is the axis-aligned box `(l, t, w, h)` in its own coordinates. */
	function shapedNode(
		id: string,
		aabb: number[],
		shape: [number, number, number, number],
	): Element {
		const g = node(id, aabb[0]!, aabb[1]!, aabb[2]!, aabb[3]!);
		const path = document.createElement('path') as unknown as Record<string, unknown> & Element;
		const [left, top, width, height] = shape;
		path.getScreenCTM = () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
		path.isPointInFill = (p: { x: number; y: number }) =>
			p.x >= left && p.x <= left + width && p.y >= top && p.y <= top + height;
		g.appendChild(path);
		return g;
	}

	it('hits a node by its drawn shape, not by a neighbour whose bounding box overlaps', () => {
		// Two turned nodes whose bounding boxes overlap at (100, 100), though only
		// "b" is actually drawn there.
		const root = document.createElement('div');
		root.append(
			shapedNode('a', [0, 0, 200, 200], [0, 0, 40, 40]),
			shapedNode('b', [80, 80, 200, 200], [90, 90, 40, 40]),
		);
		expect(smartArtNodeAtPoint(root, 100, 100)?.getAttribute('data-smartart-node-id')).toBe('b');
		expect(smartArtNodeAtPoint(root, 20, 20)?.getAttribute('data-smartart-node-id')).toBe('a');
		expect(smartArtNodeAtPoint(root, 60, 60)).toBeNull();
	});
});

describe('routeEditLayerPointerToNodes', () => {
	function layerWithNode(): { layer: HTMLElement; node: Element } {
		const layer = document.createElement('div');
		const node = document.createElement('div');
		node.setAttribute('data-smartart-node-id', 'n1');
		node.getBoundingClientRect = () =>
			({ left: 0, top: 0, width: 100, height: 100, right: 100, bottom: 100 }) as DOMRect;
		layer.append(node);
		document.body.append(layer);
		return { layer, node };
	}

	it('forwards a double-click on the node fill to the node, once', () => {
		const { layer, node } = layerWithNode();
		const seen: string[] = [];
		node.addEventListener('dblclick', () => seen.push('node'));
		const stop = routeEditLayerPointerToNodes(layer);
		// The pointer is on the layer, not on the node's own (unpainted) group.
		layer.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, clientX: 10, clientY: 10 }));
		expect(seen).toStrictEqual(['node']);
		stop();
		layer.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, clientX: 10, clientY: 10 }));
		expect(seen).toStrictEqual(['node']);
	});

	it('forwards a hover only when the hovered node changes', () => {
		const { layer, node } = layerWithNode();
		let overs = 0;
		node.addEventListener('mouseover', () => (overs += 1));
		routeEditLayerPointerToNodes(layer);
		for (const x of [5, 10, 20]) {
			layer.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: 5 }));
		}
		expect(overs).toBe(1);
	});

	it("leaves the layer's own controls alone", () => {
		const { layer, node } = layerWithNode();
		const button = document.createElement('button');
		layer.append(button);
		let dbl = 0;
		node.addEventListener('dblclick', () => (dbl += 1));
		routeEditLayerPointerToNodes(layer);
		button.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, clientX: 10, clientY: 10 }));
		expect(dbl).toBe(0);
	});

	it('ignores a role=group ancestor of the layer', () => {
		const { layer, node } = layerWithNode();
		const wrapper = document.createElement('div');
		wrapper.setAttribute('role', 'group');
		document.body.append(wrapper);
		wrapper.append(layer);
		let dbl = 0;
		node.addEventListener('dblclick', () => (dbl += 1));
		routeEditLayerPointerToNodes(layer);
		layer.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, clientX: 10, clientY: 10 }));
		expect(dbl).toBe(1);
	});
});

describe('routeEditLayerPointerToNodes hover lifecycle', () => {
	function setup(): { layer: HTMLElement; node: Element } {
		const layer = document.createElement('div');
		const node = document.createElement('div');
		node.setAttribute('data-smartart-node-id', 'n1');
		node.getBoundingClientRect = () =>
			({ left: 0, top: 0, width: 100, height: 100, right: 100, bottom: 100 }) as DOMRect;
		layer.append(node);
		document.body.append(layer);
		return { layer, node };
	}
	const move = (layer: HTMLElement, x: number, y: number): void => {
		layer.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: y }));
	};

	it('sends enter and leave as the pointer moves onto and off a node', () => {
		const { layer, node } = setup();
		const seen: string[] = [];
		for (const type of ['mouseover', 'mouseenter', 'mouseout', 'mouseleave']) {
			node.addEventListener(type, () => seen.push(type));
		}
		routeEditLayerPointerToNodes(layer);
		move(layer, 5, 5);
		expect(seen).toStrictEqual(['mouseover', 'mouseenter']);
		// Off every node (the gap), then back on: left, then re-entered.
		move(layer, 500, 500);
		expect(seen.slice(2)).toStrictEqual(['mouseout', 'mouseleave']);
		move(layer, 5, 5);
		expect(seen.slice(4)).toStrictEqual(['mouseover', 'mouseenter']);
	});

	it('leaves the node when the pointer leaves the layer', () => {
		const { layer, node } = setup();
		let left = 0;
		node.addEventListener('mouseleave', () => (left += 1));
		routeEditLayerPointerToNodes(layer);
		move(layer, 5, 5);
		layer.dispatchEvent(new MouseEvent('mouseleave'));
		expect(left).toBe(1);
	});
});
