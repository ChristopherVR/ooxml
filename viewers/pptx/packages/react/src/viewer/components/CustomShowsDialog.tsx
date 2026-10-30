/**
 * CustomShowsDialog
 *
 * Slide Show > Custom Slide Show > Custom Shows. New, Copy and Remove for the shows and, for the
 * selected one, PowerPoint's two lists (slides in the presentation, slides in the custom show)
 * with Add, Remove and Move Up/Down. Every decision is the shared `createCustomShowsController`;
 * this component only draws it. The host renders it only while open, so each open starts from a
 * fresh draft of the deck's shows.
 */
import { createCustomShowsController } from 'ooxml-ui/pptx';
import type { CustomShowsSlideRow } from 'ooxml-ui/pptx';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import React, { useMemo, useState } from 'react';

import { useController } from '../hooks/useController';
import { useTranslation } from '../i18n/use-translation';
import { DialogFooter } from './DialogFooter';
import { WebSelect } from './WebControls';

export interface CustomShowsDialogProps {
	customShows: PptxCustomShow[];
	slides: PptxSlide[];
	/** The show playback is restricted to, or null for the whole deck. */
	activeCustomShowId: string | null;
	/** Commit the shows; `clearedSlides` is set when jumps into removed shows were cleared. */
	onSave: (shows: PptxCustomShow[], clearedSlides: PptxSlide[] | undefined) => void;
	/** Restrict (or, with null, stop restricting) the slide show to one show. */
	onSetActive: (id: string | null) => void;
	/** Start the slide show (after the dialog committed and selected the show). */
	onRun: () => void;
	onClose: () => void;
}

const LIST_CLASS =
	'm-0 p-0 max-h-56 overflow-auto list-none rounded border border-border text-[12px]';
const BUTTON_CLASS =
	'px-2 py-1 rounded border border-border bg-muted hover:bg-accent text-[12px] disabled:opacity-50';

function Rows(props: {
	label: string;
	rows: readonly CustomShowsSlideRow[];
	keyOf: (row: CustomShowsSlideRow) => string | number;
	onToggle: (row: CustomShowsSlideRow) => void;
	onDoubleClick?: (row: CustomShowsSlideRow) => void;
	empty?: string;
}): React.ReactElement {
	return (
		<div className='min-w-0'>
			<strong className='block text-[11px] mb-1'>{props.label}</strong>
			<div
				role='listbox'
				aria-multiselectable='true'
				aria-label={props.label}
				className={LIST_CLASS}
			>
				{props.rows.map((row) => (
					<div
						key={props.keyOf(row)}
						role='option'
						tabIndex={0}
						aria-selected={row.selected}
						className={
							row.selected
								? 'px-2 py-1 cursor-pointer bg-primary text-white'
								: 'px-2 py-1 cursor-pointer'
						}
						onClick={() => props.onToggle(row)}
						onDoubleClick={() => props.onDoubleClick?.(row)}
						onKeyDown={(event) => {
							if (event.key === ' ') {
								event.preventDefault();
								props.onToggle(row);
							}
						}}
					>
						{row.number}. {row.title}
					</div>
				))}
				{props.rows.length === 0 && props.empty ? (
					<div className='px-2 py-1 opacity-70'>{props.empty}</div>
				) : null}
			</div>
		</div>
	);
}

export function CustomShowsDialog({
	customShows,
	slides,
	activeCustomShowId,
	onSave,
	onSetActive,
	onRun,
	onClose,
}: CustomShowsDialogProps): React.ReactElement {
	const { t } = useTranslation();
	const controller = useMemo(
		() => createCustomShowsController(customShows, activeCustomShowId),
		// A fresh draft per open: the host mounts this only while the dialog is open.
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[],
	);
	// Subscribing re-renders on every controller change; the view is cheap to rebuild.
	useController(controller);
	const view = controller.view(slides);
	const [active, setActive] = useState<string>(activeCustomShowId ?? '');
	const pickerValue = view.shows.some(({ id }) => id === active) ? active : '';

	const commit = (): PptxCustomShow[] => {
		const result = controller.result(slides);
		onSave(result.shows, result.slides);
		// A show removed in this session must not leave playback pinned to an id nothing resolves.
		onSetActive(result.shows.some(({ id }) => id === active) ? active : null);
		return result.shows;
	};

	const handleFooter = (id: string): void => {
		if (id === 'ok') {
			commit();
			onClose();
		} else if (id === 'show') {
			const selectedId = controller.getSnapshot().selectedId;
			const show = commit().find((candidate) => candidate.id === selectedId);
			if (show) {
				onSetActive(show.id);
				onRun();
				onClose();
			}
		} else {
			onClose();
		}
	};

	return (
		<>
			<button
				type='button'
				style={{ zIndex: 1200 }}
				className='fixed inset-0 bg-black/50'
				aria-label={t('pptx.customShows.closeDialog')}
				onClick={onClose}
			/>
			<div
				style={{ zIndex: 1201 }}
				className='fixed inset-0 flex items-center justify-center pointer-events-none'
			>
				<div
					role='dialog'
					aria-modal='true'
					aria-label={t('pptx.customShows.title')}
					className='pointer-events-auto flex flex-col w-[820px] max-w-[96vw] max-h-[90vh] rounded-xl border border-border bg-background shadow-2xl max-md:fixed max-md:inset-x-0 max-md:bottom-0 max-md:top-auto max-md:w-full max-md:max-w-none max-md:max-h-[88dvh] max-md:rounded-t-2xl max-md:rounded-b-none max-md:border-x-0 max-md:border-b-0'
				>
					<div className='flex items-center justify-between px-5 py-3 border-b border-border shrink-0'>
						<h2 className='text-sm font-semibold text-foreground'>{t('pptx.customShows.title')}</h2>
						<button
							type='button'
							onClick={onClose}
							className='text-muted-foreground hover:text-foreground text-lg leading-none'
							aria-label={t('pptx.customShows.close')}
						>
							&times;
						</button>
					</div>

					<div className='flex-1 px-5 py-4 space-y-3 text-[12px] text-foreground overflow-y-auto'>
						<label className='flex items-center gap-2'>
							<span>{t('pptx.customShows.selectCustomShow')}</span>
							<WebSelect
								value={pickerValue}
								onChange={(event) => {
									setActive(event.target.value);
									onSetActive(event.target.value || null);
								}}
								className='text-[11px] flex-1 min-w-0'
								aria-label={t('pptx.customShows.selectCustomShow')}
							>
								<option value=''>{t('pptx.customShows.allSlides')}</option>
								{view.shows.map((show) => (
									<option key={show.id} value={show.id}>
										{show.name}
									</option>
								))}
							</WebSelect>
						</label>

						<div className='grid gap-3 md:grid-cols-[190px_1fr]'>
							<div className='flex flex-col gap-1.5'>
								<div
									role='listbox'
									aria-label={t('pptx.customShows.savedShows')}
									className={LIST_CLASS}
								>
									{view.shows.map((show) => (
										<div
											key={show.id}
											role='option'
											tabIndex={0}
											aria-selected={show.selected}
											className={
												show.selected
													? 'px-2 py-1 cursor-pointer bg-primary text-white'
													: 'px-2 py-1 cursor-pointer'
											}
											onClick={() => controller.select(show.id)}
											onKeyDown={(event) => {
												if (event.key === 'Enter') {
													controller.select(show.id);
												}
											}}
										>
											{show.name}
										</div>
									))}
									{view.shows.length === 0 ? (
										<div className='px-2 py-1 opacity-70'>{t('pptx.customShows.empty')}</div>
									) : null}
								</div>
								<button
									type='button'
									className={BUTTON_CLASS}
									onClick={() => controller.create(t('pptx.customShows.newDefaultName'))}
								>
									{t('pptx.customShows.createNew')}
								</button>
								<button
									type='button'
									className={BUTTON_CLASS}
									disabled={!view.canEdit}
									onClick={() => controller.copySelected()}
								>
									{t('pptx.customShows.copy')}
								</button>
								<button
									type='button'
									className={BUTTON_CLASS}
									disabled={!view.canEdit}
									onClick={() => controller.removeSelected()}
								>
									{t('pptx.customShows.remove')}
								</button>
							</div>

							{view.selected ? (
								<div className='flex flex-col gap-2 min-w-0'>
									<label className='flex items-center gap-2'>
										<span>{t('pptx.customShows.renameNameLabel')}</span>
										<input
											key={view.selected.id}
											type='text'
											defaultValue={view.selected.name}
											aria-label={t('pptx.customShows.renameNameLabel')}
											className='flex-1 min-w-0 rounded border border-border bg-muted px-2 py-1'
											onBlur={(event) => controller.rename(event.target.value)}
											onKeyDown={(event) => {
												if (event.key === 'Enter') {
													controller.rename(event.currentTarget.value);
												}
											}}
										/>
									</label>
									<div className='grid gap-2 md:grid-cols-[1fr_auto_1fr_auto]'>
										<Rows
											label={t('pptx.customShows.slidesInPresentation')}
											rows={view.presentationSlides}
											keyOf={(row) => row.rId}
											onToggle={(row) => controller.toggleLeft(row.rId)}
											onDoubleClick={(row) => controller.addSlide(row.rId)}
										/>
										<div className='flex flex-col justify-center gap-1.5'>
											<button
												type='button'
												className={BUTTON_CLASS}
												disabled={!view.canAdd}
												onClick={() => controller.addSelected(slides)}
											>
												{t('pptx.customShows.add')}
											</button>
											<button
												type='button'
												className={BUTTON_CLASS}
												disabled={!view.canRemoveSlides}
												onClick={() => controller.removeSelectedSlides()}
											>
												{t('pptx.customShows.remove')}
											</button>
										</div>
										<Rows
											label={t('pptx.customShows.slidesInShow')}
											rows={view.showSlides}
											keyOf={(row) => row.position ?? 0}
											onToggle={(row) => controller.toggleRight(row.position ?? 0)}
											empty={t('pptx.customShows.noSlidesYet')}
										/>
										<div className='flex flex-col justify-center gap-1.5'>
											<button
												type='button'
												className={BUTTON_CLASS}
												disabled={!view.canMoveUp}
												onClick={() => controller.moveUp()}
											>
												{t('pptx.customShows.moveUp')}
											</button>
											<button
												type='button'
												className={BUTTON_CLASS}
												disabled={!view.canMoveDown}
												onClick={() => controller.moveDown()}
											>
												{t('pptx.customShows.moveDown')}
											</button>
										</div>
									</div>
								</div>
							) : null}
						</div>
					</div>

					<div className='px-5 py-3 border-t border-border shrink-0'>
						<DialogFooter
							actions={[
								{ id: 'show', label: t('pptx.customShows.showButton') },
								{ id: 'cancel', label: t('pptx.common.cancel') },
								{ id: 'ok', label: t('pptx.common.ok'), variant: 'primary' },
							]}
							onAction={handleFooter}
						/>
					</div>
				</div>
			</div>
		</>
	);
}
