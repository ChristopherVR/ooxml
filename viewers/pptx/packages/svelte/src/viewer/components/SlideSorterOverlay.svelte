<script lang="ts">
	/** Slide previews over the shared `createSlideSorterController`; this file only draws it. */
	import ChevronLeft from '@lucide/svelte/icons/chevron-left';
	import ChevronRight from '@lucide/svelte/icons/chevron-right';
	import X from '@lucide/svelte/icons/x';
	import { createSlideSorterController, HIDDEN_SLIDE_DIM_OPACITY, HIDDEN_SLIDE_LABEL_KEY, HIDDEN_SLIDE_SLASH_GRADIENT, hiddenSlideCue, slideClipboardLabels, slideSorterView } from 'ooxml-ui/pptx';
	import type { CanvasSize, SlidePasteRequest } from 'ooxml-ui/pptx';
	import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

	import { useTranslator } from '../../i18n/context';
	import { controllerState } from '../state/controller-state';
	import SlideSorterContextMenu from './SlideSorterContextMenu.svelte';
	import SlideStage from './SlideStage.svelte';

	const { slides, clipboardSourceSlides, clipboardTemplates, canvasSize, mediaDataUrls, current, canEdit = false, onselect, onmove, ondelete, onduplicate, onpaste, oncopysystem, onpastesystem, ontogglehidden, onclose }: { slides: PptxSlide[]; clipboardSourceSlides?: PptxSlide[]; clipboardTemplates?: Record<string, PptxElement[]>; canvasSize: CanvasSize; mediaDataUrls: Map<string, string>; current: number; canEdit?: boolean; onselect: (index: number) => void; onmove: (from: number, to: number) => void; ondelete?: (index: number) => void; onduplicate?: (index: number) => void; onpaste?: (request: SlidePasteRequest) => void; oncopysystem?: (indexes: number[]) => Promise<void>; onpastesystem?: (insertAt: number) => Promise<void>; ontogglehidden?: (index: number) => void; onclose: () => void } = $props();
	const t = useTranslator();

	const controller = createSlideSorterController({
		slides: () => slides,
		activeIndex: () => current,
		canEdit: () => canEdit,
		clipboardSourceSlides: () => clipboardSourceSlides,
		clipboardTemplates: () => clipboardTemplates,
		onOperation: (operation, indexes) => {
			for (const index of indexes) {
				if (operation === 'delete') {
					ondelete?.(index);
				} else if (operation === 'duplicate') {
					onduplicate?.(index);
				} else {
					ontogglehidden?.(index);
				}
			}
		},
		onPaste: (request) => onpaste?.(request),
		onMove: (from, to) => onmove(from, to),
		onClose: () => onclose(),
	});
	const snapshot = controllerState(controller);
	const view = $derived(slideSorterView(snapshot.current, slides));
	const sorter = $derived(snapshot.current.sorter);
	const scale = $derived((180 * view.zoomScale) / canvasSize.width);

	function oncontextmenu(event: MouseEvent, index: number): void {
		if (controller.openMenu(index, event.clientX, event.clientY)) {
			event.preventDefault();
		}
	}
	function ondragstart(event: DragEvent, index: number): void {
		event.dataTransfer?.setData('text/plain', String(index));
		if (event.dataTransfer) {
			event.dataTransfer.effectAllowed = 'move';
		}
		controller.dragStart(index);
	}
	function ondragover(event: DragEvent, index: number): void {
		event.preventDefault();
		controller.dragOver(index);
	}
	function ondrop(event: DragEvent, index: number): void {
		event.preventDefault();
		controller.drop(index, Number.parseInt(event.dataTransfer?.getData('text/plain') ?? '', 10));
	}
	function onkeydown(event: KeyboardEvent): void {
		controller.handleKeydown(event);
	}
	let clipboardControl = $state<import('ooxml-ui/pptx').PptxUiSlideClipboardElement | undefined>(undefined);
	$effect(() => {
		if (!clipboardControl) {
			return;
		}
		clipboardControl.labels = slideClipboardLabels(t);
		clipboardControl.copyAction = oncopysystem ? () => oncopysystem(controller.selectedIndexes()) : undefined;
		clipboardControl.pasteAction = canEdit && onpastesystem ? () => onpastesystem(controller.pasteInsertAt()) : undefined;
	});
</script>
<svelte:window {onkeydown} />
<div class="overlay"><header><pptx-ui-slide-clipboard bind:this={clipboardControl}></pptx-ui-slide-clipboard><h2>{t('pptx.view.slideSorter')}</h2><button aria-label={t('pptx.slideSorter.close')} onclick={onclose}><X size={16} aria-hidden="true" /></button></header><main style={`grid-template-columns: repeat(${view.gridColumns}, minmax(0, 1fr))`}>{#each slides as slide,index}{@const cue = hiddenSlideCue(slide.hidden, 'sorter', index)}<article data-pptx-chrome="sorter-tile" data-pptx-touch-press="menu dblclick" class:active={view.selectedIds.has(slide.id)} class:drop-target={snapshot.current.dragOverIndex === index} data-pptx-selected={view.selectedIds.has(slide.id)} data-pptx-slide-hidden={cue.marker} draggable={canEdit ? 'true' : 'false'} oncontextmenu={(event) => oncontextmenu(event, index)} ondragstart={(event) => ondragstart(event, index)} ondragover={(event) => ondragover(event, index)} ondragleave={() => controller.dragLeave()} ondrop={(event) => ondrop(event, index)} ondragend={() => controller.dragEnd()}><button class="preview" style={`width: ${180 * view.zoomScale}px; height: ${canvasSize.height * scale}px; opacity: ${cue.hidden ? HIDDEN_SLIDE_DIM_OPACITY : 1}`} aria-label={t('pptx.compare.slideNumber', { number: index + 1 })} aria-describedby={cue.labelId} onclick={(event) => controller.select(index, event)} ondblclick={() => { onselect(index); onclose(); }}><SlideStage {slide} {canvasSize} {mediaDataUrls} {scale} presenting={false} /></button><footer><span class="meta"><span class="num" style={cue.hidden ? `background-image: ${HIDDEN_SLIDE_SLASH_GRADIENT}` : undefined}>{index+1}</span>{#if cue.hidden}<span class="hidden-label" id={cue.labelId}>{t(HIDDEN_SLIDE_LABEL_KEY)}</span>{/if}</span><button aria-label={t('pptx.animations.moveEarlier')} disabled={!canEdit || index===0} onclick={() => controller.move(index,index-1)}><ChevronLeft size={14} aria-hidden="true" /></button><button aria-label={t('pptx.animations.moveLater')} disabled={!canEdit || index===slides.length-1} onclick={() => controller.move(index,index+1)}><ChevronRight size={14} aria-hidden="true" /></button></footer></article>{/each}</main><label>{t('pptx.slideSorter.zoom')} <input type="range" min="50" max="200" step="10" value={sorter.zoom} oninput={(event) => controller.setZoom(Number(event.currentTarget.value))} aria-label={t('pptx.slideSorter.zoom')} />{sorter.zoom}%</label>{#if snapshot.current.menu}<SlideSorterContextMenu x={snapshot.current.menu.x} y={snapshot.current.menu.y} context={view.menuContext} onrun={(id) => controller.run(id)} onclose={() => controller.closeMenu()} />{/if}</div>
<style>.overlay{position:absolute;inset:0;z-index:70;overflow:auto;background:var(--pptx-background, var(--office-background, #ffffff))}header{position:sticky;z-index:2;top:0;display:flex;align-items:center;justify-content:space-between;padding:12px 18px;border-bottom:1px solid var(--pptx-border, var(--office-border, #d1d5db));background:var(--pptx-card, var(--office-surface, #f3f4f6))}h2{margin:0;font-size:14px}header button{display:inline-flex;align-items:center;justify-content:center;border:0;background:transparent;color:inherit}main{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:18px;padding:24px}article{display:grid;justify-content:center;gap:6px;padding:9px;border:2px solid transparent;border-radius:8px;background:var(--pptx-card, var(--office-surface, #f3f4f6))}article.active{border-color:var(--pptx-primary, var(--office-accent, #2563eb))}article.drop-target{outline:2px dashed var(--pptx-primary, var(--office-accent, #2563eb));outline-offset:2px}.preview{width:180px;height:calc(180px * 9 / 16);overflow:hidden;border:0;padding:0;background:#fff;text-align:left}article footer{display:flex;align-items:center;gap:5px;font-size:11px}article footer .meta{display:flex;flex:1;align-items:center;gap:5px}article footer .num{display:inline-block;padding:0 3px}article footer .hidden-label{font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280))}article footer button{display:inline-flex;align-items:center;justify-content:center;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:4px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}article footer button:disabled{opacity:.4}</style>
