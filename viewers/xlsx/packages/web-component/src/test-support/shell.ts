// Shell test support (imported by *.test.ts only): a real EditorCore on a detached host with an
// open shadow root, so registry, ribbon and chrome tests run against the shell's own context.
import { createWorkbook, type Workbook } from '@christophervr/xlsx-core';
import type { Command } from '../commands';
import { EditorCore } from '../editor-core';

export interface ShellFixture {
	core: EditorCore;
	host: HTMLElement;
	events: { type: string; detail: unknown }[];
}

/** A core with `workbook` loaded (a blank one by default) and every public event recorded. */
export function shellFixture(workbook: Workbook | null = createWorkbook()): ShellFixture {
	const host = document.createElement('div');
	document.body.append(host);
	host.attachShadow({ mode: 'open' });
	const events: { type: string; detail: unknown }[] = [];
	for (const type of [
		'ribbon-action',
		'selection-change',
		'sheet-change',
		'dirty-change',
		'workbook-change',
		'workbook-error',
		'readonly-change',
		'file-command',
		'ribbon-customize',
	])
		host.addEventListener(type, (event) =>
			events.push({ type, detail: (event as CustomEvent).detail }),
		);
	const core = new EditorCore(host);
	if (workbook) core.setWorkbook(workbook);
	return { core, host, events };
}

/** A command that records its runs; `state` drives enabled/checked/value. */
export function spyCommand(
	id: string,
	extra: Partial<Command> = {},
): Command & { runs: unknown[] } {
	const runs: unknown[] = [];
	return {
		id,
		label: id,
		icon: 'bold',
		run: (_ctx, arg) => {
			runs.push(arg);
		},
		runs,
		...extra,
	};
}

/** Lets queued microtasks (debounced renders) run. */
export const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
