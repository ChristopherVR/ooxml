import type { RibbonControl, RibbonGroup } from 'ooxml-core/xlsx/ui';
import { commandButton, el, renderControl, type RenderScope } from './controls';
import { ribbonIcon } from './icons';

const isSmall = (control: RibbonControl): boolean =>
	control.kind === 'button' ||
	control.kind === 'toggle' ||
	control.kind === 'split' ||
	control.kind === 'menu'
		? control.size !== 'large'
		: control.kind === 'color' || control.kind === 'select';

/**
 * Lays out a run of small controls the way Excel's groups look: separators break rows (Font,
 * Alignment), a leading drop-down gets a row of its own (Number), anything else fills columns of
 * three (Clipboard's Cut/Copy/Format Painter, Editing's AutoSum/Fill/Clear).
 */
function layoutSmall(scope: RenderScope, run: RibbonControl[]): HTMLElement[] {
	const { doc } = scope;
	const render = (controls: RibbonControl[]) => {
		const box = el(doc, 'div', 'ribbon-stack');
		for (const control of controls) {
			const node = renderControl(scope, control, true);
			if (node) box.append(node);
		}
		return box;
	};
	let lines: RibbonControl[][] | undefined;
	if (run.some((control) => control.kind === 'separator')) {
		lines = [[]];
		for (const control of run)
			if (control.kind === 'separator') lines.push([]);
			else lines[lines.length - 1]!.push(control);
	} else if (run[0]?.kind === 'select' && run.some((control) => control.kind !== 'select')) {
		const lead = run.findIndex((control) => control.kind !== 'select');
		lines = [run.slice(0, lead), run.slice(lead)];
	}
	if (lines) {
		const column = el(doc, 'div', 'ribbon-stack');
		for (const line of lines.filter((items) => items.length)) {
			const row = render(line);
			if (row.childElementCount) column.append(row);
		}
		return column.childElementCount ? [column] : [];
	}
	const columns: HTMLElement[] = [];
	for (let at = 0; at < run.length; at += 3) {
		const box = render(run.slice(at, at + 3));
		if (box.childElementCount) columns.push(box);
	}
	return columns;
}

export function renderGroup(scope: RenderScope, group: RibbonGroup): HTMLElement | null {
	const { ctx, doc } = scope;
	const node = el(doc, 'div', 'ribbon-group');
	const caption = ctx.t(group.label);
	node.dataset.group = group.id;
	node.dataset.label = group.label;
	node.dataset.caption = caption;
	node.setAttribute('role', 'group');
	node.setAttribute('aria-label', caption);
	let run: RibbonControl[] = [];
	const flush = () => {
		node.append(...layoutSmall(scope, run));
		run = [];
	};
	for (const control of group.controls) {
		if (isSmall(control) || control.kind === 'separator') {
			run.push(control);
			continue;
		}
		flush();
		const rendered = renderControl(scope, control, false);
		if (rendered) node.append(rendered);
	}
	flush();
	const launcher = group.launcher ? ctx.commands.get(group.launcher) : undefined;
	if (launcher) {
		const button = commandButton(scope, launcher, { size: 'small' });
		button.className = 'ribbon-launcher';
		button.replaceChildren(ribbonIcon(doc, 'launcher', 10));
		button.dataset.launcher = '';
		node.append(button);
	}
	return node.childElementCount ? node : null;
}
