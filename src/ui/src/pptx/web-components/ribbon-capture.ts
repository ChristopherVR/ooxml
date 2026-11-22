import { createImageElement } from 'ooxml-core/pptx';
import type { PptxElement } from 'ooxml-core/pptx';
import {
	openDisplayCapture,
	previewDisplay,
	stopDisplayCapture,
	startDisplayRecording,
	type CapturedDisplayRecording,
} from '../../form/display-capture';
import { mediaElementFromFile } from '../render/building-blocks/media-element-from-file';
import type { CanvasSize } from '../types';
import {
	captureScreenshot,
	createDisplayRegionPicker,
	type DisplayRegion,
} from '../../form/display-clipping';
import { cropDisplayStream } from '../../form/display-recording-region';
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
	const options = doc.createElement('div');
	function checkbox(caption: string): OfficeUiCheckbox {
		const field = doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox;
		field.checked = true;
		const label = doc.createElement('label');
		label.style.display = 'block';
		field.setAttribute('aria-label', caption);
		label.append(field, ' ', caption);
		options.append(label);
		return field;
	}
	const audio = checkbox(text('pptx.insert.captureAudio', 'Include shared audio'));
	const pointer = checkbox(text('pptx.insert.recordPointer', 'Record pointer'));
	const select = doc.createElement('button');
	select.type = 'button';
	select.textContent = text('pptx.insert.selectArea', 'Select Area');
	const previewPanel = doc.createElement('div');
	const controls = doc.createElement('div');
	controls.style.cssText = 'display:flex;gap:8px;margin-block:12px';
	const pause = doc.createElement('button');
	pause.type = 'button';
	pause.hidden = true;
	const status = doc.createElement('div');
	status.setAttribute('role', 'alert');
	controls.append(select, pause);
	body.append(options, controls, previewPanel, status);
	return new Promise<PptxElement | undefined>((resolve) => {
		const abort = new AbortController();
		let recording: CapturedDisplayRecording | undefined;
		let busy = false;
		let source: MediaStream | undefined;
		let preview: Awaited<ReturnType<typeof previewDisplay>> | undefined;
		let region: DisplayRegion | undefined;
		function clearSelection(): void {
			source?.getVideoTracks()[0]?.removeEventListener('ended', clearSelection);
			if (source) stopDisplayCapture(source);
			source = undefined;
			preview?.release();
			preview = undefined;
			region = undefined;
			previewPanel.replaceChildren();
			sync();
		}
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
				clearSelection();
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		const sync = () => {
			const active = recording?.recorder.state;
			audio.disabled = busy || !!source;
			pointer.disabled = audio.disabled;
			select.hidden = busy || !!source;
			previewPanel.hidden = busy;
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
				(!busy && (!preview || !!region)) || (active !== undefined && active !== 'inactive'),
			);
		};
		select.addEventListener('click', () => void selectArea());
		async function selectArea(): Promise<void> {
			busy = true;
			status.textContent = '';
			sync();
			try {
				source = await openDisplayCapture(audio.checked, abort.signal, pointer.checked);
				if (!source) return;
				const capture = source;
				const next = await previewDisplay(doc, capture);
				if (abort.signal.aborted || capture.getVideoTracks()[0]?.readyState !== 'live') {
					next.release();
					clearSelection();
					return;
				}
				preview = next;
				capture.getVideoTracks()[0]?.addEventListener('ended', clearSelection, { once: true });
				const picker = createDisplayRegionPicker(
					preview.image,
					{
						heading: select.textContent ?? '',
						x: text('pptx.inspector.x', 'X'),
						y: text('pptx.inspector.y', 'Y'),
						width: text('pptx.inspector.w', 'Width'),
						height: text('pptx.inspector.h', 'Height'),
					},
					(value) => {
						region = value;
						sync();
					},
					64,
				);
				previewPanel.append(picker.element);
				picker.update();
				previewPanel.querySelector('input')?.focus();
			} catch {
				clearSelection();
				showFailure();
			} finally {
				busy = false;
				sync();
				previewPanel.querySelector('input')?.focus();
			}
		}
		function showFailure(): void {
			status.textContent = text('pptx.graphic.insertFailed', 'Could not insert the selected item.');
		}
		async function start(): Promise<void> {
			busy = true;
			status.textContent = '';
			sync();
			try {
				source ??= await openDisplayCapture(audio.checked, abort.signal, pointer.checked);
				if (!source) return;
				source.getVideoTracks()[0]?.removeEventListener('ended', clearSelection);
				if (preview && region) {
					const crop = cropDisplayStream(preview, source, region);
					recording = startDisplayRecording(crop.stream, source, crop.release);
				} else {
					recording = startDisplayRecording(source);
				}
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
				if (abort.signal.aborted || !host.isConnected) return;
				resolve({ ...media, name: heading });
				frame.dialog.close();
			} catch {
				showFailure();
			} finally {
				recording = undefined;
				clearSelection();
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
