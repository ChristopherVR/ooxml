import { CAMERA_CLIP_EXTENSION } from 'ooxml-core/pptx/editor/render/slide-operations';
import { showOfficeServiceDialog } from '../../dialog/office-service-dialog';
import { pickVideoRecordingFormat, VIDEO_MIME_CANDIDATES } from '../export/video-format';
import type { Translator } from '../i18n/translator';
import { mediaElementFromFile } from '../render/building-blocks/media-element-from-file';
import type { RecordTarget } from '../render/ribbon-record-commands';

const CAMERA_FORMATS = [
	...VIDEO_MIME_CANDIDATES.filter((type) => type.startsWith('video/mp4')),
	...VIDEO_MIME_CANDIDATES.filter((type) => type.startsWith('video/webm')),
];

/** Attach camera footage as an embedded video after the user chooses Stop and Insert. */
export function recordCameraClip(
	host: HTMLElement,
	initial: RecordTarget,
	target: () => RecordTarget | undefined,
	t: Translator,
): void {
	const doc = host.ownerDocument;
	const view = doc.defaultView;
	const mediaDevices = view?.navigator.mediaDevices;
	const Recorder = view?.MediaRecorder;
	const format =
		Recorder && pickVideoRecordingFormat(CAMERA_FORMATS, Recorder.isTypeSupported?.bind(Recorder));
	const heading = t('pptx.record.cameraClip');
	const dialog = showOfficeServiceDialog(host, heading, t('pptx.common.close'));
	const status = doc.createElement('p');
	status.setAttribute('role', 'status');
	if (!mediaDevices?.getUserMedia || !Recorder || !format || !view) {
		status.textContent = t('pptx.cameo.unavailable');
		dialog.append(status);
		dialog.show();
		return;
	}

	const preview = doc.createElement('video');
	preview.autoplay = true;
	preview.muted = true;
	preview.playsInline = true;
	preview.style.maxWidth = '100%';
	preview.hidden = true;
	const mic = doc.createElement('input');
	mic.type = 'checkbox';
	const micLabel = doc.createElement('label');
	micLabel.append(mic, ` ${t('pptx.record.includeMic')}`);
	const control = doc.createElement('button');
	control.type = 'button';
	control.textContent = t('pptx.titleBar.record');
	dialog.append(micLabel, doc.createTextNode(' '), control, preview, status);

	let stream: MediaStream | undefined;
	let recorder: MediaRecorder | undefined;
	let canceled = false;
	let saveRequested = false;
	const current = () => {
		const live = target();
		return !canceled &&
			dialog.open &&
			host.isConnected &&
			live?.slides === initial.slides &&
			live.slideIndex === initial.slideIndex
			? live
			: undefined;
	};
	const release = () => {
		for (const track of stream?.getTracks() ?? []) track.stop();
		stream = undefined;
		preview.srcObject = null;
	};
	const cancel = () => {
		canceled = true;
		observer.disconnect();
		view.removeEventListener('pagehide', cancel);
		const active = recorder;
		recorder = undefined;
		if (active && active.state !== 'inactive') active.stop();
		release();
	};
	const observer = new MutationObserver(() => {
		if (!dialog.isConnected) cancel();
	});
	observer.observe(doc, { childList: true, subtree: true });
	const root = host.getRootNode();
	if (root !== doc) observer.observe(root, { childList: true, subtree: true });
	view.addEventListener('pagehide', cancel, { once: true });
	dialog.addEventListener('office-dialog-close', cancel, { once: true });
	const fail = () => {
		recorder = undefined;
		release();
		preview.hidden = true;
		control.disabled = false;
		mic.disabled = false;
		control.textContent = t('pptx.titleBar.record');
		status.textContent = t('pptx.record.cameraClipFailed');
	};
	const save = async (chunks: Blob[]) => {
		release();
		if (!current()) {
			dialog.close();
			return;
		}
		const file = new File(chunks, `camera.${format.container}`, { type: format.blobType });
		try {
			const clip = file.size ? await mediaElementFromFile(file, initial.canvasSize, doc) : null;
			if (clip?.type !== 'media') {
				fail();
				return;
			}
			current()?.insert({ ...clip, name: heading, extensionXml: [CAMERA_CLIP_EXTENSION] });
			dialog.close();
		} catch {
			if (current()) fail();
		}
	};
	control.addEventListener('click', async () => {
		control.disabled = true;
		if (recorder?.state === 'recording') {
			saveRequested = true;
			recorder.stop();
			return;
		}
		mic.disabled = true;
		status.textContent = '';
		saveRequested = false;
		try {
			stream = await mediaDevices.getUserMedia({ video: true, audio: mic.checked });
			if (!current()) {
				release();
				dialog.close();
				return;
			}
			preview.hidden = false;
			preview.srcObject = stream;
			const active = new Recorder(stream, { mimeType: format.mimeType });
			const chunks: Blob[] = [];
			recorder = active;
			active.addEventListener('dataavailable', ({ data }) => {
				if (data.size) chunks.push(data);
			});
			active.addEventListener(
				'error',
				() => {
					if (recorder === active) fail();
				},
				{ once: true },
			);
			active.addEventListener(
				'stop',
				() => {
					if (recorder !== active) return;
					if (saveRequested) void save(chunks);
					else fail();
				},
				{ once: true },
			);
			active.start();
			control.disabled = false;
			control.textContent = t('pptx.record.stopAndInsert');
			status.textContent = t('pptx.export.recording');
		} catch {
			if (current()) fail();
		}
	});
	dialog.show();
}
