import { expect, it } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { ViewerController } from './controller';

const parser = async () => structuredClone(demoDocument);
const editor = async () => {
	throw new Error('not used');
};

it('waits for the before-show hook with loading up, then shows the drawing', async () => {
	const controller = new ViewerController(parser, () => {}, editor);
	let release: () => void = () => {};
	const seen: VisioDocument[] = [];
	controller.setBeforeShow((document) => {
		seen.push(document);
		return new Promise<void>((resolve) => (release = resolve));
	});
	const loading = controller.load(new Uint8Array([1]));
	await new Promise((resolve) => setTimeout(resolve, 0));
	expect(seen).toHaveLength(1);
	expect(controller.state.loading).toBe(true);
	expect(controller.state.document).toBeNull();
	release();
	await loading;
	expect(controller.state.loading).toBe(false);
	expect(controller.state.document).toBe(seen[0]);
	controller.destroy();
});

it('ignores a rejecting hook and lets a newer load win over a waiting one', async () => {
	const controller = new ViewerController(parser, () => {}, editor);
	controller.setBeforeShow(() => Promise.reject(new Error('offline')));
	await controller.load(new Uint8Array([1]));
	expect(controller.state.document).not.toBeNull();
	let release: () => void = () => {};
	controller.setBeforeShow(() => new Promise<void>((resolve) => (release = resolve)));
	const first = controller.load(new Uint8Array([1]));
	await new Promise((resolve) => setTimeout(resolve, 0));
	const shown = controller.state.document;
	controller.cancelLoad();
	release();
	await first;
	expect(controller.state.document).toBe(shown);
	// No hook result: no extra wait.
	controller.setBeforeShow(() => undefined);
	await controller.load(new Uint8Array([1]));
	expect(controller.state.document).not.toBe(shown);
	controller.destroy();
});
