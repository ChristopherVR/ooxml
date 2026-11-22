/** Display capture stays local and starts only inside a user gesture. */
export function canCaptureDisplay(): boolean {
	return typeof globalThis.navigator?.mediaDevices?.getDisplayMedia === 'function';
}

export async function openDisplayCapture(
	audio: boolean,
	signal?: AbortSignal,
	pointer?: boolean,
): Promise<MediaStream | undefined> {
	if (signal?.aborted) return undefined;
	try {
		const stream = await navigator.mediaDevices.getDisplayMedia({
			video: {
				width: { ideal: 1920 },
				frameRate: { ideal: 30 },
				...(pointer === undefined ? {} : { cursor: pointer ? 'always' : 'never' }),
			} as MediaTrackConstraints,
			audio,
		});
		if (!signal?.aborted) return stream;
		stopDisplayCapture(stream);
		return undefined;
	} catch (error) {
		if (error instanceof DOMException && ['AbortError', 'NotAllowedError'].includes(error.name)) {
			return undefined;
		}
		throw error;
	}
}

export function stopDisplayCapture(stream: MediaStream): void {
	for (const track of stream.getTracks()) track.stop();
}

export async function captureDisplayFrame(doc: Document): Promise<HTMLCanvasElement | undefined> {
	const stream = await openDisplayCapture(false);
	if (!stream) return undefined;
	try {
		const preview = await previewDisplay(doc, stream);
		preview.release();
		return preview.image;
	} finally {
		stopDisplayCapture(stream);
	}
}

export async function previewDisplay(doc: Document, stream: MediaStream) {
	const video = doc.createElement('video');
	video.muted = true;
	video.playsInline = true;
	video.srcObject = stream;
	try {
		await playCapturePreview(video, stream);
		const { videoWidth, videoHeight } = video;
		if (!(videoWidth > 0 && videoHeight > 0)) throw new Error('No screen frame available');
		const scale = Math.min(1, 4096 / Math.max(videoWidth, videoHeight));
		const canvas = doc.createElement('canvas');
		canvas.width = Math.max(1, Math.round(videoWidth * scale));
		canvas.height = Math.max(1, Math.round(videoHeight * scale));
		const context = canvas.getContext('2d');
		if (!context) throw new Error('Could not render screen frame');
		context.drawImage(video, 0, 0, canvas.width, canvas.height);
		return {
			image: canvas,
			video,
			release: () => {
				video.srcObject = null;
			},
		};
	} catch (error) {
		video.srcObject = null;
		throw error;
	}
}

function playCapturePreview(video: HTMLVideoElement, stream: MediaStream): Promise<void> {
	return new Promise((resolve, reject) => {
		const track = stream.getVideoTracks()[0];
		const failure = () => finish(new Error('No screen frame available'));
		const timeout = setTimeout(failure, 8000);
		function finish(error?: unknown): void {
			clearTimeout(timeout);
			track?.removeEventListener('ended', failure);
			if (error) reject(error);
			else resolve();
		}
		track?.addEventListener('ended', failure, { once: true });
		void video.play().then(() => finish(), finish);
	});
}

export interface CapturedDisplayRecording {
	readonly recorder: MediaRecorder;
	readonly result: Promise<Blob | undefined>;
	cancel(): void;
}

export function startDisplayRecording(
	stream: MediaStream,
	source: MediaStream = stream,
	release?: () => void,
): CapturedDisplayRecording {
	const cleanup = () => {
		release?.();
		stopDisplayCapture(stream);
		if (source !== stream) stopDisplayCapture(source);
	};
	const mimeType = ['video/mp4', 'video/webm;codecs=vp9,opus', 'video/webm'].find((type) =>
		MediaRecorder.isTypeSupported(type),
	);
	let recorder: MediaRecorder;
	try {
		recorder = new MediaRecorder(stream, {
			...(mimeType ? { mimeType } : {}),
			videoBitsPerSecond: 3_000_000,
		});
		recorder.start(250);
	} catch (error) {
		cleanup();
		throw error;
	}
	let discard = false;
	const chunks: Blob[] = [];
	let bytes = 0;
	const stop = () => {
		if (recorder.state !== 'inactive') recorder.stop();
	};
	const tracks = source.getVideoTracks();
	const result = new Promise<Blob | undefined>((resolve, reject) => {
		const finish = () => {
			for (const track of tracks) track.removeEventListener('ended', stop);
			cleanup();
		};
		for (const track of tracks) track.addEventListener('ended', stop, { once: true });
		recorder.addEventListener('dataavailable', ({ data }) => {
			if (!data.size || discard) return;
			bytes += data.size;
			if (bytes > 64 * 1024 * 1024) {
				discard = true;
				finish();
				stop();
				reject(new Error('The screen recording is too large'));
				return;
			}
			chunks.push(data);
		});
		recorder.addEventListener(
			'stop',
			() => {
				finish();
				resolve(
					discard || !bytes
						? undefined
						: new Blob(chunks, { type: recorder.mimeType || mimeType || 'video/webm' }),
				);
			},
			{ once: true },
		);
		recorder.addEventListener(
			'error',
			() => {
				finish();
				reject(new Error('Could not record this screen'));
			},
			{ once: true },
		);
	});
	return {
		recorder,
		result,
		cancel() {
			discard = true;
			stop();
			cleanup();
		},
	};
}
