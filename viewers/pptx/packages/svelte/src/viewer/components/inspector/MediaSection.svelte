<script lang="ts">
	import type { MediaBookmark, MediaCaptionTrack, MediaPptxElement, PptxElement } from 'pptx-viewer-core';
	import {
		MEDIA_TRIM_SECTIONS,
		mediaTrimResetPatch,
		mediaTrimResetState,
		MEDIA_PLAYBACK_SECTIONS,
		linkedContentUrl,
		formatMediaBookmarkTime,
		parseMediaBookmarkTime,
	} from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../i18n/context';

	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';
	import MediaTrimTimeline from './MediaTrimTimeline.svelte';
	import { getLinkedContentGetter } from '../../state/linked-content';
	import { resolveMediaPreviewUrl } from './media-preview';

	const { editor, mediaDataUrls = new Map() }: { editor: EditorState; mediaDataUrls?: Map<string, string> } = $props();
	const t = useTranslator();
	const resetTrimState = $derived(mediaTrimResetState(editor.selectedElement?.type === 'media' ? editor.selectedElement : undefined, editor.editable));
	const media = $derived(editor.selectedElement?.type === 'media' ? editor.selectedElement : undefined);
	// The preview plays a linked file only once this viewer allows linked content.
	const getLinkedContent = getLinkedContentGetter();
	const previewUrl = $derived(
		media ? linkedContentUrl(resolveMediaPreviewUrl(media, mediaDataUrls), getLinkedContent()) : undefined,
	);
	// eslint-disable-next-line prefer-const
	let preview = $state<HTMLMediaElement>();
	// eslint-disable-next-line prefer-const
	let duration = $state(0);
	let currentTime = $state(0);
	// eslint-disable-next-line prefer-const
	let playing = $state(false);
	let bookmarkLabel = $state('');
	// eslint-disable-next-line prefer-const
	let bookmarkTime = $state(formatMediaBookmarkTime(0));
	// The trim fields (`MEDIA_TRIM_SECTIONS`) are bounded by the live duration.
	const durationMs = $derived(duration * 1000);

	function patch(next: Partial<MediaPptxElement>): void {
		if (media) {
			editor.applyElementPatch(media.id, next as Partial<PptxElement>);
		}
	}
	function addBookmark(): void {
		const time = parseMediaBookmarkTime(bookmarkTime);
		if (!media || time === undefined) {
			return;
		}
		const label = bookmarkLabel.trim() || `${t('pptx.media.bookmark')} ${(media.bookmarks?.length ?? 0) + 1}`;
		patch({ bookmarks: [...(media.bookmarks ?? []), { id: `bookmark-${Date.now()}`, label, time }] });
		bookmarkLabel = '';
	}
	function updateBookmarkTime(bookmark: MediaBookmark, input: HTMLInputElement): void {
		const time = parseMediaBookmarkTime(input.value, bookmark.time) ?? bookmark.time;
		input.value = formatMediaBookmarkTime(time);
		if (time !== bookmark.time) {
			patch({ bookmarks: (media?.bookmarks ?? []).map((item) => item.id === bookmark.id ? { ...item, time } : item) });
		}
	}
	function addTrack(): void {
		if (!media) {
			return;
		}
		const count = media.captionTracks?.length ?? 0;
		patch({ captionTracks: [...(media.captionTracks ?? []), { id: `caption-${Date.now()}`, label: `${t('pptx.media.trackName')} ${count + 1}`, language: 'en', kind: 'subtitles', isDefault: count === 0 }] });
	}
	function trackPatch(index: number, next: Partial<MediaCaptionTrack>): void {
		if (media) {
			patch({ captionTracks: (media.captionTracks ?? []).map((track, i) => i === index ? { ...track, ...next } : track) });
		}
	}
	function togglePreview(): void {
		if (!preview) {
			return;
		}
		if (preview.paused) {
			void preview.play();
		} else {
			preview.pause();
		}
	}
	function seek(seconds: number): void {
		if (preview) {
			preview.currentTime = seconds;
		}
		currentTime = seconds;
	}
</script>

{#if media}<div class="section">
	{#if previewUrl}<div class="preview">{#if media.mediaType === 'video'}<video bind:this={preview} src={previewUrl} preload="metadata" onloadedmetadata={() => (duration = preview && Number.isFinite(preview.duration) ? preview.duration : 0)} ontimeupdate={() => { if (preview) currentTime = preview.currentTime; }} onplay={() => (playing = true)} onpause={() => (playing = false)}><track kind="captions" label={t('pptx.media.captions')} src="data:text/vtt;charset=utf-8,WEBVTT" />{#each media.captionTracks ?? [] as track}<track kind={track.kind === 'subtitles' ? 'subtitles' : 'captions'} label={track.label} srclang={track.language} src={track.src ?? `data:text/vtt;charset=utf-8,${encodeURIComponent(track.content ?? 'WEBVTT')}`} default={track.isDefault} />{/each}</video>{:else}<audio bind:this={preview} src={previewUrl} preload="metadata" onloadedmetadata={() => (duration = preview && Number.isFinite(preview.duration) ? preview.duration : 0)} ontimeupdate={() => { if (preview) currentTime = preview.currentTime; }} onplay={() => (playing = true)} onpause={() => (playing = false)}></audio>{/if}<button type="button" onclick={togglePreview}>{playing ? t('pptx.media.pause') : t('pptx.media.play')}</button><span>{currentTime.toFixed(1)} / {duration.toFixed(1)}s</span>{#if duration > 0}<MediaTrimTimeline {duration} startMs={media.trimStartMs ?? 0} endMs={media.trimEndMs} {currentTime} bookmarks={media.bookmarks ?? []} onchange={(trimStartMs, trimEndMs) => patch({ trimStartMs, trimEndMs })} onseek={seek} />{/if}</div>{/if}
	<InspectorFields sections={MEDIA_PLAYBACK_SECTIONS} target={media as PptxElement} disabled={!editor.editable} onpatch={(next) => patch(next as Partial<MediaPptxElement>)} />
	<InspectorFields sections={MEDIA_TRIM_SECTIONS} target={{ media, durationMs }} disabled={!editor.editable} onpatch={(next) => patch(next as Partial<MediaPptxElement>)} />
	{#if resetTrimState.visible}<button type="button" class="reset-trim" onclick={() => patch(mediaTrimResetPatch() as Partial<MediaPptxElement>)}>{t('pptx.media.resetTrim')}</button>{/if}
	<details open><summary>{t('pptx.media.bookmarks')}</summary>{#each [...(media.bookmarks ?? [])].sort((a,b)=>a.time-b.time) as bookmark}<div class="row"><input value={bookmark.label} onchange={(event) => patch({bookmarks:(media.bookmarks??[]).map((item)=>item.id===bookmark.id?{...item,label:event.currentTarget.value}:item)})} /><input type="text" value={formatMediaBookmarkTime(bookmark.time)} onchange={(event) => updateBookmarkTime(bookmark, event.currentTarget)} /><button aria-label={t('pptx.media.removeBookmark')} onclick={() => patch({bookmarks:(media.bookmarks??[]).filter((item)=>item.id!==bookmark.id)})}>×</button></div>{/each}<div class="row"><input placeholder={t('pptx.media.bookmarkLabel')} bind:value={bookmarkLabel} /><input type="text" bind:value={bookmarkTime} /><button aria-label={t('pptx.media.addBookmark')} onclick={addBookmark}>+</button></div></details>
	<details open><summary>{t('pptx.media.captionTracks')}</summary>{#each media.captionTracks ?? [] as track, index}<fieldset><div class="grid"><label>{t('pptx.media.trackLabel')}<input value={track.label} onchange={(event)=>trackPatch(index,{label:event.currentTarget.value})} /></label><label>{t('pptx.media.trackLanguage')}<input value={track.language} onchange={(event)=>trackPatch(index,{language:event.currentTarget.value})} /></label><label>{t('pptx.media.trackKind')}<pptx-ui-select aria-label={t('pptx.media.trackKind')} value={track.kind} onchange={(event)=>trackPatch(index,{kind:event.currentTarget.value as MediaCaptionTrack['kind']})}><option value="subtitles">{t('pptx.media.trackKind.subtitles')}</option><option value="captions">{t('pptx.media.captions')}</option><option value="descriptions">{t('pptx.media.trackKind.descriptions')}</option></pptx-ui-select></label><label class="inline"><pptx-ui-checkbox checked={track.isDefault ?? false} onchange={(event)=>patch({captionTracks:(media.captionTracks??[]).map((item,i)=>({...item,isDefault:event.currentTarget.checked?i===index:(i===index?false:item.isDefault)}))})} ></pptx-ui-checkbox>{t('pptx.media.trackDefault')}</label></div><label>{t('pptx.media.trackSource')}<input value={track.src ?? ''} placeholder={t('pptx.media.trackSourcePlaceholder')} onchange={(event)=>trackPatch(index,{src:event.currentTarget.value||undefined})} /></label><label>{t('pptx.media.trackInlineWebVtt')}<textarea rows="4" value={track.content ?? ''} onchange={(event)=>trackPatch(index,{content:event.currentTarget.value||undefined})}></textarea></label><button onclick={() => patch({captionTracks:(media.captionTracks??[]).filter((_item,i)=>i!==index)})}>{t('pptx.media.removeTrack')}</button></fieldset>{/each}<button onclick={addTrack}>{t('pptx.media.addCaptionTrack')}</button></details>
</div>{/if}

<style>.section{display:grid;gap:8px}.preview{display:grid;grid-template-columns:auto 1fr;gap:5px;align-items:center}.preview video,.preview audio{grid-column:1/-1;width:100%;max-height:130px;background:#000}.preview :global(.timeline),.preview :global(.times){grid-column:1/-1}.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}.inline{display:flex;align-items:center}label{display:grid;gap:3px;color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-size:10px}input,select,textarea{min-width:0;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit}input,select{height:25px}.row{display:grid;grid-template-columns:1fr 70px 28px;gap:4px;margin-top:5px}details{border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));padding-top:7px}summary{cursor:pointer;font-weight:600}fieldset{display:grid;gap:6px;margin:6px 0;padding:6px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:6px}button{border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;padding:4px 7px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}</style>
