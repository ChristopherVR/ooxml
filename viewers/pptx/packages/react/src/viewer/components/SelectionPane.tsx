import {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	selectionPaneRows,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
/**
 * Selection Pane: lists all elements on the active slide with
 * visibility toggles, rename-on-double-click, and drag-to-reorder. The row
 * labels, the rename box and the drag are the shared selection pane controller.
 */
import React, { useState, useCallback, useRef } from 'react';
import { LuEye, LuEyeOff, LuGripVertical } from 'react-icons/lu';

import { useController } from '../hooks/useController';
import { useTranslation } from '../i18n/use-translation';
import { cn } from '../utils';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface SelectionPaneProps {
	slides: PptxSlide[];
	activeSlideIndex: number;
	selectedElementId: string | null;
	selectedElementIds: string[];
	canEdit: boolean;
	setSelectedElementId: (id: string | null) => void;
	setSelectedElementIds: (ids: string[]) => void;
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	markDirty: () => void;
	onClose: () => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function SelectionPane({
	slides,
	activeSlideIndex,
	selectedElementId,
	selectedElementIds,
	canEdit,
	setSelectedElementId,
	setSelectedElementIds,
	setSlides,
	markDirty,
	onClose,
}: SelectionPaneProps): React.ReactElement {
	const { t } = useTranslation();
	const activeSlide = slides[activeSlideIndex] as PptxSlide | undefined;
	const elements = activeSlide?.elements ?? [];

	const inputRef = useRef<HTMLInputElement>(null);
	const latest = useRef({ activeSlideIndex, canEdit, setSlides, markDirty });
	latest.current = { activeSlideIndex, canEdit, setSlides, markDirty };
	/** Rewrite the active slide's element list (rename or reorder), marking the deck dirty. */
	const updateElements = useCallback(
		(update: (elements: PptxElement[]) => PptxElement[] | null) => {
			const { activeSlideIndex: index } = latest.current;
			latest.current.setSlides((prevSlides) => {
				const slide = prevSlides[index];
				const next = slide ? update([...(slide.elements ?? [])]) : null;
				if (!slide || !next) {
					return prevSlides;
				}
				const nextSlides = [...prevSlides];
				nextSlides[index] = { ...slide, elements: next };
				return nextSlides;
			});
			latest.current.markDirty();
		},
		[],
	);
	const [pane] = useState(() =>
		createSelectionPaneController({
			canEdit: () => latest.current.canEdit,
			// A rename round-trips through save as cNvPr/@name, and so does a clear: `''` is
			// written as `name=""` (`@name` is required), while `undefined` means "no opinion".
			onRename: (id, name) =>
				updateElements((list) => {
					const idx = list.findIndex((e) => e.id === id);
					if (idx === -1) {
						return null;
					}
					list[idx] = { ...list[idx], name } as PptxElement;
					return list;
				}),
			onReorder: (from, to) =>
				updateElements((list) => {
					const [moved] = list.splice(from, 1);
					list.splice(to, 0, moved);
					return list;
				}),
		}),
	);
	const paneState = useController(pane);

	const handleSelect = useCallback(
		(elementId: string) => {
			setSelectedElementId(elementId);
			setSelectedElementIds([elementId]);
		},
		[setSelectedElementId, setSelectedElementIds],
	);

	const handleToggleVisibility = useCallback(
		(elementId: string) => {
			setSlides((prevSlides) => {
				const nextSlides = [...prevSlides];
				const slide = nextSlides[activeSlideIndex];
				if (!slide) {
					return prevSlides;
				}
				const nextElements = [...(slide.elements ?? [])];
				const idx = nextElements.findIndex((e) => e.id === elementId);
				if (idx === -1) {
					return prevSlides;
				}
				nextElements[idx] = {
					...nextElements[idx],
					hidden: !nextElements[idx].hidden,
				} as PptxElement;
				nextSlides[activeSlideIndex] = { ...slide, elements: nextElements };
				return nextSlides;
			});
			markDirty();
		},
		[activeSlideIndex, setSlides, markDirty],
	);

	const handleDoubleClick = useCallback(
		(element: PptxElement, displayName: string) => {
			if (pane.startRename(element.id, displayName)) {
				setTimeout(() => inputRef.current?.select(), 0);
			}
		},
		[pane],
	);

	// Top-most element first.
	const rows = selectionPaneRows(elements);

	return (
		<div
			data-pptx-selection-pane
			className='flex flex-col h-full bg-popover border-l border-border w-56'
		>
			<div className='flex items-center justify-between px-3 py-2 border-b border-border'>
				<span className='text-xs font-medium text-foreground'>{t('pptx.selectionPane.title')}</span>
				<button
					type='button'
					onClick={onClose}
					className='text-muted-foreground hover:text-foreground text-xs'
					title={t('pptx.selectionPane.close')}
				>
					&times;
				</button>
			</div>
			<div className='flex-1 overflow-y-auto py-1'>
				{rows.length === 0 ? (
					<div className='px-3 py-4 text-xs text-muted-foreground italic'>
						{t('pptx.selectionPane.empty')}
					</div>
				) : (
					rows.map(({ element, index: realIndex, label }) => {
						const isSelected =
							selectedElementId === element.id || selectedElementIds.includes(element.id);
						const displayName = formatSelectionPaneLabel(label, t);
						const isEditing = paneState.renamingId === element.id;

						return (
							<div
								key={element.id}
								draggable={canEdit}
								onDragStart={() => pane.dragStart(realIndex)}
								onDragOver={(e) => {
									e.preventDefault();
									pane.dragOver(realIndex);
								}}
								onDrop={() => pane.drop(realIndex)}
								onDragEnd={pane.dragEnd}
								className={cn(
									'flex items-center gap-1.5 px-2 py-1 text-xs cursor-pointer transition-colors',
									isSelected ? 'bg-primary/30 text-primary' : 'hover:bg-muted text-foreground',
									paneState.dragOverIndex === realIndex ? 'border-t-2 border-primary' : '',
								)}
								onClick={() => handleSelect(element.id)}
								onDoubleClick={() => handleDoubleClick(element, displayName)}
							>
								{canEdit && (
									<LuGripVertical className='w-3 h-3 text-muted-foreground flex-shrink-0 cursor-grab' />
								)}
								<span data-pptx-selection-name className='flex-1 truncate'>
									{isEditing ? (
										<input
											ref={inputRef}
											type='text'
											aria-label={t('pptx.selectionPane.renameElement')}
											value={paneState.renameDraft}
											onChange={(e) => pane.setRenameDraft(e.target.value)}
											onBlur={() => pane.commitRename()}
											onKeyDown={(e) => {
												if (e.key === 'Enter') {
													pane.commitRename();
												}
												if (e.key === 'Escape') {
													pane.cancelRename();
												}
											}}
											className='w-full bg-muted text-xs px-1 py-0.5 rounded border border-border outline-none'
											onClick={(e) => e.stopPropagation()}
										/>
									) : (
										displayName
									)}
								</span>
								<button
									type='button'
									className='text-muted-foreground hover:text-foreground flex-shrink-0'
									title={
										element.hidden ? t('pptx.selectionPane.show') : t('pptx.selectionPane.hide')
									}
									onClick={(e) => {
										e.stopPropagation();
										handleToggleVisibility(element.id);
									}}
								>
									{element.hidden ? (
										<LuEyeOff className='w-3.5 h-3.5' />
									) : (
										<LuEye className='w-3.5 h-3.5 opacity-50' />
									)}
								</button>
							</div>
						);
					})
				)}
			</div>
		</div>
	);
}
