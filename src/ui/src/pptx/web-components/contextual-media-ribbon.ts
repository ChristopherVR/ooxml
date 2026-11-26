import { html, nothing } from 'lit';
import type { MediaPptxElement } from 'ooxml-core/pptx';
import { generateElementId } from '../render/element-clipboard';
import { canInteractWithElement } from '../render/element-locks';
import { MEDIA_PLAYBACK_SECTION } from '../render/inspector/media-playback-fields';
import type { ContextualControls, ContextualFields } from './contextual-ribbon-fields';

/** The live canvas node. The slide rail can contain another copy of the same id. */
export function ribbonMediaNode(host: HTMLElement, id: string): HTMLMediaElement | undefined {
	const root = host.closest('[data-pptx-editor-chrome]') ?? host.getRootNode();
	if (!(root instanceof Element || root instanceof Document || root instanceof ShadowRoot)) return;
	const selector = `[data-element-id="${CSS.escape(id)}"][data-pptx-element="true"]`;
	for (const wrapper of root.querySelectorAll(selector)) {
		const media = wrapper.matches('audio,video') ? wrapper : wrapper.querySelector('audio,video');
		if (media instanceof HTMLMediaElement && !media.closest('[data-pptx-chrome="slides"]'))
			return media;
	}
	return undefined;
}

export function renderMediaControls(
	ctx: ContextualControls,
	fields: ContextualFields,
	media: MediaPptxElement,
) {
	const { text, patch, number, toggle, select, color, command } = fields;
	const options = MEDIA_PLAYBACK_SECTION.fields;
	const volumeChoices = [
		['low', text('pptx.mediaRibbon.volumeLow', 'Low'), 0.33],
		['medium', text('pptx.mediaRibbon.volumeMedium', 'Medium'), 0.5],
		['high', text('pptx.mediaRibbon.volumeHigh', 'High'), 0.8],
		['mute', text('pptx.mediaRibbon.volumeMute', 'Mute'), 0],
	] as const;
	const volume = media.volume ?? 1;
	const volumeChoice =
		media.muted || volume === 0
			? 'mute'
			: volume < 0.415
				? 'low'
				: volume < 0.65
					? 'medium'
					: 'high';
	const fieldLabel = (id: string, fallback: string) =>
		text(options.find((field) => field.id === id)?.labelKey ?? '', fallback);
	const local =
		!media.isLinked && !media.mediaMissing && Boolean(media.mediaPath || media.mediaData);
	const writable = !media.rawXml || Boolean(media.rawXml['p:nvPicPr']);
	const unsupported = text(
		'pptx.mediaRibbon.unsupported',
		'This command requires embedded media in a picture.',
	);
	const sourceRequired = text(
		'pptx.mediaRibbon.localRequired',
		'This command requires an embedded media clip.',
	);
	const pictureRequired = writable ? undefined : unsupported;
	const clipRequired = pictureRequired ?? (local ? undefined : sourceRequired);
	const dialog = (kind: 'trim' | 'remove' | 'poster' | 'current' | 'reset' | 'crop') => {
		const signal = ctx.begin();
		void import('./contextual-media-dialogs').then(
			(module) => {
				if (!signal.aborted) module.runMediaRibbonCommand(kind, ctx, fields, media, signal);
			},
			() => {
				if (!signal.aborted) ctx.close();
			},
		);
	};
	const mediaFlag = (
		control: string,
		key: 'loop' | 'fullScreen' | 'hideWhenNotPlaying' | 'playAcrossSlides' | 'rewindAfterPlaying',
		label: string,
		reason?: string,
	) =>
		toggle(
			control,
			label,
			media[key] === true,
			(enabled) => patch({ [key]: enabled }),
			reason ?? pictureRequired,
		);
	switch (ctx.group) {
		case 'playback.bookmarks':
			return html`${command(
				'add',
				text('pptx.media.addBookmark', 'Add Bookmark'),
				() => {
					const time = ribbonMediaNode(ctx.host, media.id)?.currentTime;
					const bookmarks = media.bookmarks ?? [];
					if (
						time === undefined ||
						!Number.isFinite(time) ||
						time < 0 ||
						time > 4294967 ||
						bookmarks.length >= 256 ||
						bookmarks.some((mark) => Math.round(mark.time * 1000) === Math.round(time * 1000))
					)
						return;
					const label = text('pptx.media.bookmark', 'Bookmark');
					const used = new Set(bookmarks.map((mark) => mark.label));
					let index = bookmarks.length + 1;
					while (used.has(`${label} ${index}`)) index += 1;
					patch({
						bookmarks: [
							...bookmarks,
							{ id: generateElementId(), time, label: `${label} ${index}` },
						],
					});
				},
				clipRequired,
			)}
			${command(
				'remove',
				text('pptx.mediaRibbon.removeBookmark', 'Remove Bookmark'),
				() => dialog('remove'),
				pictureRequired ??
					(media.bookmarks?.length
						? undefined
						: text('pptx.mediaRibbon.noBookmarks', 'This clip has no bookmarks.')),
			)}`;
		case 'playback.editing':
			return html`${command('trim', text('pptx.media.trim', 'Trim'), () => dialog('trim'), clipRequired)}
			${number('fadeIn', fieldLabel('fadeInDuration', 'Fade In (s)'), media.fadeInDuration ?? 0, 0, 59, 0.001, (value) => patch({ fadeInDuration: value || undefined }), clipRequired)}
			${number('fadeOut', fieldLabel('fadeOutDuration', 'Fade Out (s)'), media.fadeOutDuration ?? 0, 0, 59, 0.001, (value) => patch({ fadeOutDuration: value || undefined }), clipRequired)}`;
		case 'playback.options':
			return html`${select(
				'volume',
				fieldLabel('volume', 'Volume'),
				volumeChoice,
				volumeChoices.map(([id, label]) => [id, label]),
				(value) => {
					const choice = volumeChoices.find(([id]) => id === value);
					if (choice)
						patch(value === 'mute' ? { muted: true } : { volume: choice[2], muted: false });
				},
				pictureRequired,
			)}
			${select(
				'start',
				fieldLabel('start', 'Start'),
				media.startMode ?? (media.autoPlay ? 'auto' : 'click'),
				[
					['sequence', text('pptx.mediaRibbon.inSequence', 'In Click Sequence')],
					['auto', text('pptx.media.startAutomatically', 'Automatically')],
					['click', text('pptx.mediaRibbon.whenClicked', 'When Clicked On')],
				],
				(value) => {
					if (value !== 'auto' && value !== 'click' && value !== 'sequence') return;
					patch({
						startMode: value,
						autoPlay: value === 'auto',
					});
				},
				pictureRequired,
			)}
			${mediaFlag(
				'fullScreen',
				'fullScreen',
				fieldLabel('fullScreen', 'Play Full Screen'),
				media.mediaType === 'audio'
					? text('pptx.mediaRibbon.videoOnly', 'This command is for video.')
					: undefined,
			)}
			${mediaFlag('hide', 'hideWhenNotPlaying', fieldLabel('hideWhenNotPlaying', 'Hide While Not Playing'))}
			${mediaFlag('loop', 'loop', text('pptx.mediaRibbon.loop', 'Loop until Stopped'))}
			${mediaFlag('rewind', 'rewindAfterPlaying', text('pptx.mediaRibbon.rewind', 'Rewind after Playing'))}
			${mediaFlag(
				'across',
				'playAcrossSlides',
				fieldLabel('playAcrossSlides', 'Play Across Slides'),
				media.mediaType !== 'audio'
					? text('pptx.mediaRibbon.audioOnly', 'This command is for audio.')
					: undefined,
			)}`;
		case 'videoFormat.poster':
			return html`${command('current', text('pptx.mediaRibbon.currentFrame', 'Current Frame'), () => dialog('current'), clipRequired)}
			${command('file', text('pptx.mediaRibbon.fromFile', 'Image from File'), () => dialog('poster'), pictureRequired)}
			${command('reset', text('pptx.mediaRibbon.resetPoster', 'Reset Poster Frame'), () => dialog('reset'), clipRequired)}`;
		case 'audioFormat.pictureStyles':
		case 'videoFormat.pictureStyles':
			return html`${color(
				'border',
				text('pptx.pictureFormat.border', 'Picture Border'),
				media.shapeStyle?.strokeColor ?? '#000000',
				(strokeColor) =>
					patch({
						shapeStyle: {
							...media.shapeStyle,
							strokeColor,
							strokeColorXml: undefined,
							strokeFillMode: 'solid',
							strokeWidth: media.shapeStyle?.strokeWidth || 1,
						},
					}),
				pictureRequired,
			)}`;
		case 'audioFormat.arrange':
		case 'videoFormat.arrange':
			return command(
				'rotateRight',
				text('pptx.pictureFormat.rotateRight', 'Rotate Right 90°'),
				() => patch({ rotation: ((media.rotation ?? 0) + 90) % 360 }),
				canInteractWithElement(media, 'rotate') ? undefined : fields.unavailable,
			);
		case 'audioFormat.size':
		case 'videoFormat.size': {
			const resize = canInteractWithElement(media, 'resize') ? undefined : fields.unavailable;
			return html`${number('width', text('pptx.arrange.width', 'Width'), media.width, 1, 9600, 'any', (width) => patch({ width }), resize)}
			${number('height', text('pptx.arrange.height', 'Height'), media.height, 1, 9600, 'any', (height) => patch({ height }), resize)}
			${
				ctx.group === 'videoFormat.size'
					? command(
							'crop',
							text('pptx.image.crop', 'Crop'),
							() => dialog('crop'),
							pictureRequired ??
								(canInteractWithElement(media, 'crop') ? undefined : fields.unavailable),
						)
					: nothing
			}`;
		}
		default:
			return nothing;
	}
}
