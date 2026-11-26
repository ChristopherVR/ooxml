import { MAX_PICTURE_EDIT_BYTES, readPictureRasterHeader } from 'ooxml-core/pptx';
import type { MediaPptxElement } from 'ooxml-core/pptx';
import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import { mediaTrimEndAbsoluteMs, mediaTrimEndMsFromAbsoluteMs } from '../render/media-trim-range';
import { ribbonMediaNode } from './contextual-media-ribbon';
import type { ContextualControls, ContextualFields } from './contextual-ribbon-fields';
import { pictureDataUrl } from './picture-worker-client';

type MediaCommand = 'trim' | 'remove' | 'poster' | 'current' | 'reset' | 'crop';

export function runMediaRibbonCommand(
	kind: MediaCommand,
	ctx: ContextualControls,
	fields: ContextualFields,
	media: MediaPptxElement,
	signal: AbortSignal,
): void {
	const { host } = ctx;
	const close = () => {
		if (!signal.aborted) ctx.close();
	};
	const { text } = fields;
	const doc = host.ownerDocument;
	const body = doc.createElement('div');
	body.style.cssText = 'display:grid;gap:12px';
	const status = doc.createElement('div');
	status.setAttribute('role', 'status');
	body.append(status);
	const video = ribbonMediaNode(host, media.id);
	const patch = (value: Partial<MediaPptxElement>) => {
		if (!signal.aborted) fields.patch(value);
		close();
	};
	const title = {
		trim: text('pptx.media.trim', 'Trim'),
		remove: text('pptx.mediaRibbon.removeBookmark', 'Remove Bookmark'),
		poster: text('pptx.mediaRibbon.fromFile', 'Image from File'),
		current: text('pptx.mediaRibbon.currentFrame', 'Current Frame'),
		reset: text('pptx.mediaRibbon.resetPoster', 'Reset Poster Frame'),
		crop: text('pptx.image.crop', 'Crop'),
	}[kind];
	let apply = () => {};
	const frame = createInsertDialogFrame(host, body, () => apply(), close);
	frame.dialog.dataset.mediaRibbonDialog = kind;
	const focus = host.shadowRoot?.activeElement ?? doc.activeElement;
	signal.addEventListener(
		'abort',
		() => {
			frame.dialog.close();
			frame.dialog.remove();
			if (focus instanceof HTMLElement) queueMicrotask(() => focus.focus());
		},
		{ once: true },
	);
	const sync = (valid: boolean) =>
		frame.sync(
			title,
			text('pptx.common.close', 'Close'),
			text('pptx.common.cancel', 'Cancel'),
			text('pptx.common.apply', 'Apply'),
			valid,
		);
	const field = (label: string, input: HTMLElement) => {
		const wrapper = doc.createElement('label');
		wrapper.style.cssText = 'display:flex;gap:8px;align-items:center';
		wrapper.append(label, input);
		body.append(wrapper);
	};
	const number = (label: string, value: number, max: number) => {
		const input = doc.createElement('input');
		input.type = 'number';
		input.min = '0';
		input.max = String(max);
		input.step = 'any';
		input.required = true;
		input.value = String(value);
		field(label, input);
		return input;
	};
	const fail = () => {
		status.textContent = text(
			'pptx.mediaRibbon.sourceUnavailable',
			'The media preview or image is unavailable.',
		);
		sync(false);
	};
	sync(false);
	frame.dialog.show();

	switch (kind) {
		case 'trim': {
			const duration = video?.duration || media.metadata?.duration || 0;
			if (!Number.isFinite(duration) || duration <= 0 || duration > 4294967) return fail();
			const durationMs = Math.round(duration * 1000);
			const start = number(
				`${text('pptx.media.trimStart', 'Trim Start')} (s)`,
				(media.trimStartMs ?? 0) / 1000,
				duration,
			);
			const end = number(
				`${text('pptx.media.trimEnd', 'Trim End')} (s)`,
				mediaTrimEndAbsoluteMs(durationMs, media.trimEndMs ?? 0) / 1000,
				duration,
			);
			const valid = () =>
				start.validity.valid &&
				end.validity.valid &&
				Math.round(start.valueAsNumber * 1000) < Math.round(end.valueAsNumber * 1000);
			body.addEventListener('input', () => sync(valid()));
			apply = () => {
				if (!valid()) return;
				patch({
					trimStartMs: Math.round(start.valueAsNumber * 1000) || undefined,
					trimEndMs:
						mediaTrimEndMsFromAbsoluteMs(durationMs, Math.round(end.valueAsNumber * 1000)) ||
						undefined,
				});
			};
			sync(valid());
			return;
		}
		case 'remove': {
			const input = doc.createElement('pptx-ui-select');
			const label = text('pptx.media.bookmark', 'Bookmark');
			input.setAttribute('aria-label', label);
			const bookmarks = (media.bookmarks ?? []).slice(0, 256);
			for (const mark of bookmarks) {
				const option = doc.createElement('option');
				option.value = mark.id;
				option.textContent = `${mark.label} (${mark.time}s)`;
				input.append(option);
			}
			field(label, input);
			input.value = bookmarks[0]?.id ?? '';
			apply = () =>
				patch({ bookmarks: media.bookmarks?.filter((mark) => mark.id !== input.value) });
			sync(Boolean(input.value));
			return;
		}
		case 'poster': {
			const input = doc.createElement('input');
			input.type = 'file';
			input.accept = 'image/png,image/jpeg';
			field(text('pptx.pictureFormat.chooseRaster', 'Choose a PNG or JPEG'), input);
			input.addEventListener('change', () => sync(Boolean(input.files?.[0])));
			apply = () => {
				const file = input.files?.[0];
				if (!file || file.size > MAX_PICTURE_EDIT_BYTES) return fail();
				input.disabled = true;
				sync(false);
				void file
					.arrayBuffer()
					.then((bytes) => {
						signal.throwIfAborted();
						const header = readPictureRasterHeader(bytes);
						if (!header) throw new Error('Unsupported poster image');
						return pictureDataUrl(file.slice(0, file.size, header.type));
					})
					.then(
						(data) => patch({ posterFrameData: data, posterFramePath: undefined }),
						() => {
							if (signal.aborted) return;
							input.disabled = false;
							fail();
						},
					);
			};
			return;
		}
		case 'current':
		case 'reset': {
			if (
				!(video instanceof HTMLVideoElement) ||
				!video.videoWidth ||
				!video.videoHeight ||
				!/^(?:blob:|data:video\/)/u.test(video.currentSrc)
			)
				return fail();
			void captureFrame(video, kind === 'reset', signal).then(
				(data) => patch({ posterFrameData: data, posterFramePath: undefined }),
				() => {
					if (!signal.aborted) fail();
				},
			);
			return;
		}
		case 'crop': {
			const inputs = {
				cropLeft: number(
					text('pptx.image.cropLeft', 'Crop Left') + ' %',
					(media.cropLeft ?? 0) * 100,
					99,
				),
				cropRight: number(
					text('pptx.image.cropRight', 'Crop Right') + ' %',
					(media.cropRight ?? 0) * 100,
					99,
				),
				cropTop: number(
					text('pptx.image.cropTop', 'Crop Top') + ' %',
					(media.cropTop ?? 0) * 100,
					99,
				),
				cropBottom: number(
					text('pptx.image.cropBottom', 'Crop Bottom') + ' %',
					(media.cropBottom ?? 0) * 100,
					99,
				),
			};
			const valid = () =>
				Object.values(inputs).every((input) => input.validity.valid) &&
				inputs.cropLeft.valueAsNumber + inputs.cropRight.valueAsNumber < 100 &&
				inputs.cropTop.valueAsNumber + inputs.cropBottom.valueAsNumber < 100;
			body.addEventListener('input', () => sync(valid()));
			apply = () => {
				if (valid())
					patch(
						Object.fromEntries(
							Object.entries(inputs).map(([key, input]) => [key, input.valueAsNumber / 100]),
						),
					);
			};
			sync(valid());
		}
	}
}

async function captureFrame(
	video: HTMLVideoElement,
	reset: boolean,
	signal: AbortSignal,
): Promise<string> {
	// Cap the new bitmap before allocation. The original video is already decoded on the stage.
	const ratio = Math.min(1, 1280 / Math.max(video.videoWidth, video.videoHeight));
	const width = Math.max(1, Math.round(video.videoWidth * ratio));
	const height = Math.max(1, Math.round(video.videoHeight * ratio));
	const position = video.currentTime;
	const seek = reset && position !== 0;
	try {
		if (seek) {
			const deadline = AbortSignal.any([signal, AbortSignal.timeout(4000)]);
			await new Promise<void>((resolve, reject) => {
				deadline.throwIfAborted();
				video.addEventListener('seeked', () => resolve(), { once: true, signal: deadline });
				deadline.addEventListener('abort', () => reject(deadline.reason), { once: true });
				video.currentTime = 0;
			});
		}
		signal.throwIfAborted();
		const canvas = new OffscreenCanvas(width, height);
		const context = canvas.getContext('2d');
		if (!context || video.readyState < 2) throw new Error('Media frame is unavailable');
		context.drawImage(video, 0, 0, width, height);
		return pictureDataUrl(await canvas.convertToBlob({ type: 'image/png' }));
	} finally {
		if (seek) video.currentTime = position;
	}
}
