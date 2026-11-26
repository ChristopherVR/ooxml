import { setDisabledReason } from './dialog-fields';
import type { EditorCore } from './editor-core';
import { translate } from './localization';
import type { RibbonIcon } from './ribbon-icons';
import { group, stack, tool } from './ribbon-parts';

export type ReferencesCommand =
	| 'markEntry'
	| 'insertIndex'
	| 'updateIndex'
	| 'sources'
	| 'citation'
	| 'editCitation'
	| 'style'
	| 'bibliography'
	| 'worksCited'
	| 'updateBibliography'
	| 'markAuthority'
	| 'insertAuthority'
	| 'updateAuthority';

export const REFERENCE_LABELS = {
	markEntry: 'Mark entry',
	insertIndex: 'Insert index',
	updateIndex: 'Update index',
	sources: 'Manage sources',
	citation: 'Insert citation',
	editCitation: 'Edit citation',
	style: 'Citation style',
	bibliography: 'Bibliography',
	worksCited: 'Works cited',
	updateBibliography: 'Update bibliography',
	markAuthority: 'Mark citation',
	insertAuthority: 'Insert table of authorities',
	updateAuthority: 'Update table of authorities',
} as const satisfies Record<ReferencesCommand, string>;

export function buildReferences(panel: HTMLElement): void {
	const command = (key: ReferencesCommand, icon: RibbonIcon, inline = false) => {
		const button = tool(
			REFERENCE_LABELS[key],
			icon,
			{ type: 'references', key },
			{ large: !inline, inline },
		);
		button.dataset.references = key;
		if (
			[
				'markEntry',
				'insertIndex',
				'sources',
				'citation',
				'editCitation',
				'style',
				'markAuthority',
				'insertAuthority',
			].includes(key)
		)
			button.setAttribute('aria-haspopup', 'dialog');
		return button;
	};
	panel.append(
		group(
			'Index',
			command('markEntry', 'addText'),
			stack(command('insertIndex', 'toc', true), command('updateIndex', 'updateTable', true)),
		),
		group(
			'Citations & Bibliography',
			command('citation', 'crossReference'),
			stack(
				command('editCitation', 'caption', true),
				command('sources', 'table', true),
				command('style', 'changeCase', true),
			),
			stack(
				command('bibliography', 'toc', true),
				command('worksCited', 'toc', true),
				command('updateBibliography', 'updateTable', true),
			),
		),
		group(
			'Table of Authorities',
			command('markAuthority', 'addText'),
			stack(
				command('insertAuthority', 'toc', true),
				command('updateAuthority', 'updateTable', true),
			),
		),
	);
}

export function referencesReason(core: EditorCore, key: ReferencesCommand) {
	if (core.loading || !core.view?.editable || !core.host.canEditOutsideBody())
		return 'references.unavailable';
	const { $from, $to } = core.view.state.selection;
	if (key === 'markEntry' || key === 'markAuthority' || key === 'citation')
		return !$from.sameParent($to) || $from.parent.type.name !== 'paragraph'
			? 'references.paragraph'
			: undefined;
	if (key === 'editCitation') {
		const isCitation = (instr: unknown) =>
			typeof instr === 'string' && /^\s*CITATION(?:\s|$)/i.test(instr);
		const onCitation = $from
			.marks()
			.some(
				(mark) => mark.type.name === 'field' && isCitation(mark.attrs.instr) && !mark.attrs.locked,
			);
		if (!onCitation) {
			const adjacent = [$from.nodeBefore, $from.nodeAfter].some((node) =>
				node?.marks.some(
					(mark) =>
						mark.type.name === 'field' && isCitation(mark.attrs.instr) && !mark.attrs.locked,
				),
			);
			if (!adjacent) return 'references.selectCitation';
		}
	}
	return undefined;
}

export function syncReferencesRibbon(core: EditorCore): void {
	for (const button of core.shell.toolbar?.querySelectorAll<HTMLButtonElement>(
		'[data-references]',
	) ?? []) {
		const command = Object.entries(REFERENCE_LABELS).find(
			([key]) => key === button.dataset.references,
		);
		if (!command) continue;
		const [key, label] = command;
		const reason = referencesReason(core, key as ReferencesCommand);
		setDisabledReason(
			button,
			reason ? translate(core.locale, reason) : undefined,
			translate(core.locale, label),
		);
	}
}
