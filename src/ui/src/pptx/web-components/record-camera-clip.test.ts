// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { recordCameraClip } from './record-camera-clip';
import type { RecordTarget } from '../render/ribbon-record-commands';
import { mediaElementFromFile } from '../render/building-blocks/media-element-from-file';

vi.mock('../render/building-blocks/media-element-from-file', () => ({
	mediaElementFromFile: vi.fn(),
}));

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
	vi.resetAllMocks();
});

class TestRecorder extends EventTarget {
	static isTypeSupported = (type: string) => type === 'video/mp4';
	state: RecordingState = 'inactive';
	start() {
		this.state = 'recording';
	}
	stop() {
		this.state = 'inactive';
		const event = Object.assign(new Event('dataavailable'), { data: new Blob(['clip']) });
		this.dispatchEvent(event);
		this.dispatchEvent(new Event('stop'));
	}
}

function setup() {
	const stop = vi.fn();
	const stream = { getTracks: () => [{ stop }] };
	const getUserMedia = vi.fn().mockResolvedValue(stream);
	vi.stubGlobal('navigator', { mediaDevices: { getUserMedia } });
	vi.stubGlobal('MediaRecorder', TestRecorder);
	const host = document.createElement('div');
	document.body.append(host);
	const insert = vi.fn();
	const initial: RecordTarget = {
		slides: [{ id: 'slide', rId: 'rId1', slideNumber: 1, elements: [] }],
		slideIndex: 0,
		canvasSize: { width: 800, height: 600 },
		insert,
	};
	recordCameraClip(
		host,
		initial,
		() => initial,
		(key) => key,
	);
	const dialog = host.querySelector('office-ui-dialog')!;
	return { host, dialog, button: dialog.querySelector('button')!, stop, insert, getUserMedia };
}

it('requests camera permission only on Record and embeds the clip on Stop and Insert', async () => {
	const clip = {
		id: 'video',
		type: 'media' as const,
		mediaType: 'video' as const,
		x: 0,
		y: 0,
		width: 20,
		height: 20,
	};
	vi.mocked(mediaElementFromFile).mockResolvedValue(clip);
	const ui = setup();
	expect(ui.getUserMedia).not.toHaveBeenCalled();
	ui.button.click();
	await vi.waitFor(() => expect(ui.button.textContent).toBe('pptx.record.stopAndInsert'));
	expect(ui.getUserMedia).toHaveBeenCalledExactlyOnceWith({ video: true, audio: false });
	ui.button.click();
	await vi.waitFor(() =>
		expect(ui.insert).toHaveBeenCalledWith(
			expect.objectContaining({
				...clip,
				name: 'pptx.record.cameraClip',
				extensionXml: [expect.objectContaining({ uri: 'urn:ooxml:powerpoint:camera-clip' })],
			}),
		),
	);
	expect(ui.stop).toHaveBeenCalledOnce();
	expect(vi.mocked(mediaElementFromFile).mock.calls[0]?.[0].type).toBe('video/mp4');
});

it('stops capture and discards the recording when the dialog is closed', async () => {
	const ui = setup();
	ui.dialog.querySelector('input')!.checked = true;
	ui.button.click();
	await vi.waitFor(() => expect(ui.button.textContent).toBe('pptx.record.stopAndInsert'));
	expect(ui.getUserMedia).toHaveBeenCalledExactlyOnceWith({ video: true, audio: true });
	ui.dialog.close();
	expect(ui.stop).toHaveBeenCalledOnce();
	expect(ui.insert).not.toHaveBeenCalled();
	expect(mediaElementFromFile).not.toHaveBeenCalled();
});
