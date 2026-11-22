import { createImageElement } from 'ooxml-core/pptx';
import type { PptxElement } from 'ooxml-core/pptx';
import {
	openDisplayCapture,
	startDisplayRecording,
	type CapturedDisplayRecording,
} from '../../form/display-capture';
import { mediaElementFromFile } from '../render/building-blocks/media-element-from-file';
import type { CanvasSize } from '../types';
import { captureScreenshot } from '../../form/display-clipping';
import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import type { InsertText } from './ribbon-graphic-insert';
import type { OfficeUiCheckbox } from '../../form/checkbox';
import { definePptxCheckbox } from './office-aliases';

export async function pickScreenshot(
	host: HTMLElement,
	canvasSize: CanvasSize,
	clip: boolean,
	text: InsertText,
): Promise<PptxElement | undefined> {
	const selected = await captureScreenshot(
		host,
		clip
			? {
					heading: text('pptx.insert.screenClipping', 'Screen Clipping'),
					close: text('pptx.common.close', 'Close'),
					cancel: text('pptx.common.cancel', 'Cancel'),
					insert: text('pptx.common.insert', 'Insert'),
					x: text('pptx.inspector.x', 'X'),
					y: text('pptx.inspector.y', 'Y'),
					width: text('pptx.inspector.w', 'Width'),
					height: text('pptx.inspector.h', 'Height'),
				}
			: undefined,
	);
	if (!selected) return undefined;
	const scale = Math.min(1, canvasSize.width / selected.width, canvasSize.height / selected.height);
	const width = selected.width * scale;
	const height = selected.height * scale;
	const name = text('pptx.insert.screenshot', 'Screenshot');
	return {
		...createImageElement(selected.toDataURL('image/png'), {
			x: (canvasSize.width - width) / 2,
			y: (canvasSize.height - height) / 2,
			width,
			height,
			altText: name,
		}),
		name,
	};
}

export function pickScreenRecording(host: HTMLElement, canvasSize: CanvasSize, text: InsertText) {
	const doc = host.ownerDocument;
	const heading = text('pptx.insert.screenRecording', 'Screen Recording');
	const body = doc.createElement('div');
	const registry = doc.defaultView?.customElements;
	if (registry) definePptxCheckbox(registry);
	const audio = doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox;
	audio.checked = true;
	const label = doc.createElement('label');
	const audioLabel = text('pptx.insert.captureAudio', 'Include shared audio');
	audio.setAttribute('aria-label', audioLabel);
	label.append(audio, ' ', audioLabel);
	const controls = doc.createElement('div');
	controls.style.cssText = 'display:flex;gap:8px;margin-block:12px';
	const pause = doc.createElement('button');
	pause.type = 'button';
	pause.hidden = true;
	const status = doc.createElement('div');
	status.setAttribute('role', 'alert');
	controls.append(pause);
	body.append(label, controls, status);
	return new Promise<PptxElement | undefined>((resolve) => {
		const abort = new AbortController();
		let recording: CapturedDisplayRecording | undefined;
		let busy = false;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				if (recording) {
					if (recording.recorder.state !== 'inactive') recording.recorder.stop();
					return;
				}
				if (busy) return;
				void start();
			},
			() => {
				abort.abort();
				recording?.cancel();
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		const sync = () => {
			const active = recording?.recorder.state;
			audio.disabled = busy || !!recording;
			pause.hidden = !active || active === 'inactive';
			pause.textContent =
				active === 'paused'
					? text('pptx.rehearse.resume', 'Resume')
					: text('pptx.rehearse.pause', 'Pause');
			frame.sync(
				heading,
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				recording ? text('pptx.followMode.stop', 'Stop') : text('pptx.titleBar.record', 'Record'),
				!busy || (active !== undefined && active !== 'inactive'),
			);
		};
		async function start(): Promise<void> {
			busy = true;
			status.textContent = '';
			sync();
			try {
				const stream = await openDisplayCapture(audio.checked, abort.signal);
				if (!stream) return;
				recording = startDisplayRecording(stream);
				sync();
				const blob = await recording.result;
				if (!blob || abort.signal.aborted || !host.isConnected) return;
				const mimeType = blob.type.split(';')[0] ?? 'video/webm';
				const file = new File(
					[blob],
					mimeType === 'video/mp4' ? 'screen-recording.mp4' : 'screen-recording.webm',
					{ type: mimeType },
				);
				const media = await mediaElementFromFile(file, canvasSize, doc);
				if (!media) throw new Error('Could not insert screen recording');
				resolve({ ...media, name: heading });
				frame.dialog.close();
			} catch {
				status.textContent = text(
					'pptx.graphic.insertFailed',
					'Could not insert the selected item.',
				);
			} finally {
				recording = undefined;
				busy = false;
				sync();
			}
		}
		pause.addEventListener('click', () => {
			const recorder = recording?.recorder;
			if (recorder?.state === 'paused') recorder.resume();
			else if (recorder?.state === 'recording') recorder.pause();
			sync();
		});
		sync();
		frame.dialog.show();
	});
}
