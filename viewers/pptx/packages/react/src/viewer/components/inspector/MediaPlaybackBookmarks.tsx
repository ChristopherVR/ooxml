import { formatMediaBookmarkTime, MEDIA_PLAYBACK_SECTIONS } from 'ooxml-ui/pptx';
import type { MediaPptxElement, MediaBookmark, PptxElement } from 'pptx-viewer-core';
import React, { useState, useCallback } from 'react';
import { LuPlus, LuTrash2, LuBookmark } from 'react-icons/lu';

import { useTranslation } from '../../i18n/use-translation';
import { InspectorFields } from './InspectorFields';
import { CARD, HEADING, INPUT, BTN, generateBookmarkId } from './TrimTimeline';

// ==========================================================================
// Props
// ==========================================================================

interface MediaPlaybackBookmarksProps {
	element: MediaPptxElement;
	canEdit: boolean;
	currentTime: number;
	bookmarks: MediaBookmark[];
	onUpdateElement: (updates: Partial<PptxElement>) => void;
	onSeekTo: (time: number) => void;
}

// ==========================================================================
// Component
// ==========================================================================

export function MediaPlaybackBookmarks({
	element,
	canEdit,
	currentTime,
	bookmarks,
	onUpdateElement,
	onSeekTo,
}: MediaPlaybackBookmarksProps): React.ReactElement {
	const { t } = useTranslation();
	const [newBookmarkLabel, setNewBookmarkLabel] = useState('');

	const handleAddBookmark = useCallback((): void => {
		const label = newBookmarkLabel.trim() || `${t('pptx.media.bookmark')} ${bookmarks.length + 1}`;
		const newBmk: MediaBookmark = {
			id: generateBookmarkId(),
			time: currentTime,
			label,
		};
		onUpdateElement({
			bookmarks: [...bookmarks, newBmk],
		} as Partial<PptxElement>);
		setNewBookmarkLabel('');
	}, [bookmarks, currentTime, newBookmarkLabel, onUpdateElement, t]);

	const handleRemoveBookmark = useCallback(
		(bmkId: string): void => {
			onUpdateElement({
				bookmarks: bookmarks.filter((b) => b.id !== bmkId),
			} as Partial<PptxElement>);
		},
		[bookmarks, onUpdateElement],
	);

	return (
		<>
			{/* ── Playback settings ── */}
			<div className={CARD}>
				<div className={HEADING}>{t('pptx.media.playback')}</div>

				{/* Volume, start, flags, fades and speed: the shared descriptors. */}
				<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
					<InspectorFields
						sections={MEDIA_PLAYBACK_SECTIONS}
						target={element as PptxElement}
						disabled={!canEdit}
						onPatch={onUpdateElement}
					/>
				</div>
			</div>

			{/* ── Bookmarks ── */}
			<div className={CARD}>
				<div className={HEADING}>
					<LuBookmark className='inline w-3 h-3 mr-1' />
					{t('pptx.media.bookmarks')}
				</div>

				{bookmarks.length > 0 && (
					<div className='space-y-1 max-h-32 overflow-y-auto'>
						{bookmarks
							.slice()
							.sort((a, b) => a.time - b.time)
							.map((bmk) => (
								<div key={bmk.id} className='flex items-center gap-1 text-[11px] group'>
									<button
										type='button'
										className='text-primary hover:text-primary/80 truncate flex-1 text-left'
										onClick={() => onSeekTo(bmk.time)}
										title={t('pptx.media.seekToBookmark')}
									>
										{bmk.label}
									</button>
									<span className='text-muted-foreground tabular-nums text-[10px]'>
										{formatMediaBookmarkTime(bmk.time)}
									</span>
									{canEdit && (
										<button
											type='button'
											className='opacity-0 group-hover:opacity-100 focus-visible:opacity-100 pointer-coarse:opacity-100 text-red-400 hover:text-red-300 transition-opacity'
											onClick={() => handleRemoveBookmark(bmk.id)}
											title={t('pptx.common.remove')}
										>
											<LuTrash2 className='w-3 h-3' />
										</button>
									)}
								</div>
							))}
					</div>
				)}

				{canEdit && (
					<div className='flex items-center gap-1'>
						<input
							type='text'
							className={`${INPUT} text-[11px]`}
							placeholder={t('pptx.media.bookmarkLabel')}
							value={newBookmarkLabel}
							onChange={(e) => setNewBookmarkLabel(e.target.value)}
							onKeyDown={(e) => {
								if (e.key === 'Enter') {
									e.preventDefault();
									handleAddBookmark();
								}
							}}
						/>
						<button
							type='button'
							className={BTN}
							onClick={handleAddBookmark}
							title={t('pptx.media.addBookmark')}
						>
							<LuPlus className='w-3 h-3' />
						</button>
					</div>
				)}
			</div>
		</>
	);
}
