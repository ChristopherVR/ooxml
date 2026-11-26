import { isHeaderFooterDesign } from 'ooxml-core/docx';
import type { FormatDialog } from './font-dialog';
import type { InsertPartsHost } from './insert-parts-form';
import type { EditorCore } from './editor-core';
import { createInsertPartsHost } from './insert-parts-host';
import type { InsertPartsKind, RibbonAction } from './ribbon-action';
import './locales/insert-parts-translations';
import { emit } from './events';
import { translate, type LocalizationKey } from './localization';

const create = async (kind: InsertPartsKind, host: InsertPartsHost): Promise<FormatDialog> => {
	if (kind === 'fieldDialog' || kind === 'onlineVideo') {
		const more = await import('./insert-parts-more-dialogs');
		return kind === 'fieldDialog'
			? more.createFieldDialog(host)
			: more.createOnlineVideoDialog(host);
	}
	const main = await import('./insert-parts-dialogs');
	return kind === 'signatureSetup'
		? main.createSignatureSetupDialog(host)
		: kind === 'saveQuickPart'
			? main.createSaveQuickPartDialog(host)
			: main.createBuildingBlocksDialog(host);
};

const dialogs = new WeakMap<EditorCore, Map<InsertPartsKind, FormatDialog>>();

/** Every Quick Parts, Signature Line, Online Video and header or footer design command. */
export async function insertPartsAction(
	core: EditorCore,
	action: Extract<
		RibbonAction,
		{
			type:
				| 'insertParts'
				| 'quickPart'
				| 'headerFooter'
				| 'headerFooterRemove'
				| 'headerFooterEdit';
		}
	>,
	opener?: HTMLElement,
): Promise<void> {
	if (action.type === 'insertParts') return openInsertParts(core, action.kind, opener);
	if (action.type === 'quickPart') {
		const view = core.targetView();
		if (!view?.editable) return;
		// A bound control is only kept in step with the property in the document's body.
		if (action.command.kind === 'property' && view !== core.view) {
			await import('./locales/insert-parts-translations');
			emit(
				core.element,
				'document-warning',
				translate(
					core.locale,
					'Document properties can only be inserted in the body of the document.' as LocalizationKey,
				),
			);
			return;
		}
		const { runQuickPart } = await import('./quick-parts-commands');
		return runQuickPart(view, core.model, action.command, core.locale);
	}
	if (action.type === 'headerFooterEdit') return editHeaderFooter(core, action.kind);
	return changeHeaderFooterDesign(core, action);
}

/** Insert > Header or Footer > Edit: opens the story of the section holding the caret and puts the caret in it. */
function editHeaderFooter(core: EditorCore, kind: 'header' | 'footer'): void {
	const view = core.view;
	if (!view) return;
	const sectionIndex = core.pages.tools.currentSectionIndex(view, core.model);
	core.parts.showStory(
		{ kind, slot: 'default', sectionIndex } as never,
		0,
		0,
		core.shell.canvas,
		core.shell.paper,
	);
}

/** Opens one of the Quick Parts, Signature Line or Online Video dialogs, creating it on first use. */
export async function openInsertParts(
	core: EditorCore,
	kind: InsertPartsKind,
	opener?: HTMLElement,
): Promise<void> {
	await import('./locales/insert-parts-translations');
	const known = dialogs.get(core) ?? new Map<InsertPartsKind, FormatDialog>();
	dialogs.set(core, known);
	let dialog = known.get(kind);
	if (!dialog) {
		dialog = await create(kind, createInsertPartsHost(core));
		known.set(kind, dialog);
		core.shell.toolbar?.parentElement?.append(dialog.element);
	}
	core.formatDialogs.closeAll();
	dialog.setLocale(core.locale);
	dialog.open(opener);
}

/** Insert > Header or Footer > a gallery design, or Remove, for the section holding the caret. */
export async function changeHeaderFooterDesign(
	core: EditorCore,
	action: Extract<RibbonAction, { type: 'headerFooter' | 'headerFooterRemove' }>,
): Promise<void> {
	const commands = await import('./header-footer-design-commands');
	const stories = action.kind === 'header' ? 'headers' : 'footers';
	const design = action.type === 'headerFooter' ? action.design : undefined;
	if (action.type === 'headerFooter' && !(design && isHeaderFooterDesign(design))) return;
	const tools = core.pages.tools;
	const view = core.view;
	const section = (model: typeof core.model) => (view ? tools.currentSectionIndex(view, model) : 0);
	const done = core.pages.changeHeaderFooter((model) =>
		design && action.type === 'headerFooter'
			? commands.withHeaderFooterDesign(
					tools,
					model,
					stories,
					design,
					tools.newHeaderFooterId,
					section(model),
					action.slot ?? 'default',
					{
						title: translate(core.locale, 'Document title'),
						author: translate(core.locale, 'Author name' as LocalizationKey),
						date: translate(core.locale, 'Date' as LocalizationKey),
					},
				)
			: commands.withoutHeaderFooter(tools, model, stories, section(model)),
	);
	if (done) core.parts.render(core.shell.canvas, core.shell.paper, true);
}
