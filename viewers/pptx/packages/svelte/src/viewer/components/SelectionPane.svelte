<script lang="ts">
	/** The Selection Pane: rows, rename box and drag reorder from the shared selection pane controller. */
	import BringToFront from '@lucide/svelte/icons/bring-to-front';
	import Eye from '@lucide/svelte/icons/eye';
	import EyeOff from '@lucide/svelte/icons/eye-off';
	import SendToBack from '@lucide/svelte/icons/send-to-back';
	import X from '@lucide/svelte/icons/x';
	import type { PptxSlide } from 'pptx-viewer-core';
	import {
		createSelectionPaneController,
		formatSelectionPaneLabel,
		isElementHidden,
		restoreEditorKeyboardFocus,
		selectionPaneRows,
	} from 'ooxml-ui/pptx';
	import type { EditorState } from '../editor/editor-state.svelte'; import { useTranslator } from '../../i18n/context';
	import { controllerState } from '../state/controller-state';

	const { editor, onclose }: { editor: EditorState; onclose: () => void } = $props(); const t = useTranslator();
	const elements = $derived(editor.activeElements);
	const rows = $derived(selectionPaneRows(elements).map((row) => ({ ...row, name: formatSelectionPaneLabel(row.label, t) })));

	/** Flip an element's Selection Pane visibility. */
	function toggleHidden(id: string, hidden: boolean): void {
		editor.applyElementPatch(id, { hidden: !hidden });
	}

	/** Move the element at stacking position `from` to `to` on the active slide, one undo step. */
	function reorder(from: number, to: number): void {
		const index = editor.currentSlideIndex;
		const slide = editor.slides[index];
		const list = [...(slide?.elements ?? [])];
		// The pane lists the rendered layer; reorder only when it is the slide's own list.
		if (!slide || list[from]?.id !== elements[from]?.id) { return; }
		const [moved] = list.splice(from, 1);
		list.splice(to, 0, moved);
		editor.commitSlides(editor.slides.map((item, i) => (i === index ? { ...item, elements: list } : item)) as PptxSlide[]);
	}

	const pane = createSelectionPaneController({
		canEdit: () => editor.editable,
		onRename: (id, name) => editor.applyElementPatch(id, { name }),
		onReorder: reorder,
	});
	const paneState = controllerState(pane);
	let renameInput = $state<HTMLInputElement | null>(null);

	/**
	 * Give the keyboard back to the viewer root before the input goes away, so the Ctrl+Z that
	 * undoes the rename reaches the viewer (shared `restoreEditorKeyboardFocus`).
	 */
	// The box closes first, then the keyboard goes back to the root: focusing the root blurs the
	// input, and the blur's own commit must find the box already closed.
	function commitRename(): void {
		if (paneState.current.renamingId === null) { return; }
		pane.commitRename();
		restoreEditorKeyboardFocus(renameInput);
	}
	function cancelRename(): void {
		pane.cancelRename();
		restoreEditorKeyboardFocus(renameInput);
	}
	function renameKeydown(event: KeyboardEvent): void {
		if (event.key === 'Enter') { event.preventDefault(); commitRename(); }
		else if (event.key === 'Escape') { event.stopPropagation(); cancelRename(); }
	}
	/** Svelte action: focus + select the rename input the moment it mounts. */
	function autofocus(node: HTMLInputElement): void {
		renameInput = node;
		node.focus();
		node.select();
	}
	function ondragover(event: DragEvent, index: number): void {
		event.preventDefault();
		pane.dragOver(index);
	}
</script>
<aside class="pane" data-pptx-selection-pane aria-label={t('pptx.ribbon.selectionPane')}><header><h2>{t('pptx.ribbon.selectionPane')}</h2><button aria-label={t('pptx.common.close')} onclick={onclose}><X size={16} aria-hidden="true" /></button></header><div>{#each rows as { element, index, name } (element.id)}<div class="row" class:selected={editor.selection.ids.includes(element.id)} class:drop={paneState.current.dragOverIndex === index} draggable={editor.editable && paneState.current.renamingId !== element.id ? 'true' : 'false'} role="listitem" ondragstart={() => pane.dragStart(index)} ondragover={(event) => ondragover(event, index)} ondrop={() => pane.drop(index)} ondragend={() => pane.dragEnd()}>{#if paneState.current.renamingId === element.id}<input class="rename" type="text" value={paneState.current.renameDraft} oninput={(event) => pane.setRenameDraft(event.currentTarget.value)} aria-label={t('pptx.selectionPane.renameElement')} use:autofocus onkeydown={renameKeydown} onblur={commitRename} />{:else}<button class="name" onclick={() => editor.select(element.id)} ondblclick={() => pane.startRename(element.id, name)}><span data-pptx-selection-name>{name}</span><i>{element.type}</i></button>{/if}<button class="eye" title={isElementHidden(element) ? t('pptx.selectionPane.showElement') : t('pptx.selectionPane.hideElement')} aria-label={isElementHidden(element) ? t('pptx.selectionPane.showElement') : t('pptx.selectionPane.hideElement')} aria-pressed={isElementHidden(element)} onclick={() => toggleHidden(element.id, isElementHidden(element))}>{#if isElementHidden(element)}<EyeOff size={14} aria-hidden="true" />{:else}<Eye size={14} aria-hidden="true" />{/if}</button></div>{:else}<p>{t('pptx.statusBar.noSlides')}</p>{/each}</div><footer><button title={t('pptx.arrange.bringToFront')} aria-label={t('pptx.arrange.bringToFront')} disabled={!editor.selectedElementId} onclick={() => editor.reorderSelected('front')}><BringToFront size={14} aria-hidden="true" /></button><button title={t('pptx.arrange.sendToBack')} aria-label={t('pptx.arrange.sendToBack')} disabled={!editor.selectedElementId} onclick={() => editor.reorderSelected('back')}><SendToBack size={14} aria-hidden="true" /></button></footer></aside>
<style>.pane{position:absolute;z-index:45;top:0;right:0;bottom:0;width:260px;border-left:1px solid var(--pptx-border, var(--office-border, #d1d5db));background:var(--pptx-card, var(--office-surface, #f3f4f6));box-shadow:-12px 0 35px #0006}header,footer{display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--pptx-border, var(--office-border, #d1d5db))}h2{margin:0;font-size:12px}header button{display:inline-flex;align-items:center;justify-content:center;border:0;background:transparent;color:inherit}.pane>div{display:grid;gap:4px;overflow:auto;padding:8px}.row{display:flex;align-items:stretch;gap:4px;border:1px solid transparent;border-radius:5px;background:var(--pptx-muted, var(--office-surface, #f3f4f6))}.row.selected{border-color:var(--pptx-primary, var(--office-accent, #2563eb))}.row.drop{box-shadow:inset 0 2px 0 var(--pptx-primary, var(--office-accent, #2563eb))}.row button{border:0;background:transparent;color:inherit;font-size:11px}.row .name{display:flex;flex:1;justify-content:space-between;gap:8px;min-width:0;padding:8px;text-align:left}.row .name span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row .rename{flex:1;min-width:0;margin:4px;padding:3px 5px;border:1px solid var(--pptx-primary, var(--office-accent, #2563eb));border-radius:4px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit;font:inherit;font-size:11px}.row .eye{display:inline-flex;flex-shrink:0;align-items:center;justify-content:center;padding:0 8px}.row .eye[aria-pressed='false']{opacity:.5}i{color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-style:normal}footer{justify-content:flex-end;gap:5px;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-bottom:0}footer button{display:inline-flex;align-items:center;justify-content:center;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}footer button:disabled{opacity:.4}</style>
