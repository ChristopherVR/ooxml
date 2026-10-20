// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { nativeBuildingBlockContract } from '../../../scripts/native-building-block-contract';
it('solid composes native controls with reactive inputs and engine lifecycle', async () => {
	expect.hasAssertions();
	await nativeBuildingBlockContract('solid');
}, 30000);

import { render } from 'solid-js/web';
import { createSignal } from 'solid-js';
import { SpreadsheetCanvas } from './building-blocks';
import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { createWorkbook } from '../../core/src';
it('transfers canvas ownership when the native mount accessor changes', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = createBuildingBlockSession(() => {});
	const second = createBuildingBlockSession(() => {});
	first.update({ workbook: createWorkbook() });
	second.update({ workbook: createWorkbook() });
	const [mount, setMount] = createSignal(first.mount);
	const dispose = render(
		() =>
			SpreadsheetCanvas({
				get mount() {
					return mount();
				},
			}),
		host,
	);
	const old = host.querySelector('xlsx-editor')!;
	setMount(() => second.mount);
	await Promise.resolve();
	expect(old.isConnected).toBe(false);
	expect(host.querySelectorAll('xlsx-editor')).toHaveLength(1);
	expect(await first.handle.saveBytes().catch(() => null)).toBeNull();
	dispose();
	expect(host.querySelector('xlsx-editor')).toBeNull();
	host.remove();
});
