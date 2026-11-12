import { BUILDING_BLOCK_OK, buildingBlockFailure, classifyMediaType } from 'ooxml-ui/pptx';
import type { BuildingBlockResult } from 'ooxml-ui/pptx';
/**
 * insert-file-handlers: Factory functions for image and media file-picking
 * handlers used by useInsertElements.
 */
import type { PptxElement, PptxSlide, ImagePptxElement, MediaPptxElement } from 'pptx-viewer-core';

import type { CanvasSize } from '../types';
import { generateElementId } from '../utils/generate-id';

export interface FileHandlerDeps {
	activeSlide: PptxSlide | undefined;
	canvasSize: CanvasSize;
	addElement: (element: PptxElement) => void;
}

export interface FileHandlers {
	handleImageFileChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
	handleMediaFileChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
	/** Insert an audio or video file chosen by the host's own picker. */
	/**
	 * `shouldInsert` is asked once the media is ready (after a video size probe); when it says no
	 * nothing is added and the result is `stale-target`.
	 */
	insertMediaFile: (file: File, shouldInsert?: () => boolean) => Promise<BuildingBlockResult>;
}

export function createFileHandlers(deps: FileHandlerDeps): FileHandlers {
	const { activeSlide, canvasSize, addElement } = deps;

	const handleImageFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
		const file = e.target.files?.[0];
		if (!file || !activeSlide) {
			return;
		}
		e.target.value = '';
		const reader = new FileReader();
		reader.onload = () => {
			const dataUrl = reader.result as string;
			if (!dataUrl) {
				return;
			}
			const img = new Image();
			img.onload = () => {
				const MAX_W = 500,
					MAX_H = 400;
				let w = img.naturalWidth || 400,
					h = img.naturalHeight || 300;
				if (w > MAX_W || h > MAX_H) {
					const scale = Math.min(MAX_W / w, MAX_H / h);
					w = Math.round(w * scale);
					h = Math.round(h * scale);
				}
				addElement({
					id: generateElementId(),
					type: 'image',
					x: Math.round((canvasSize.width - w) / 2),
					y: Math.round((canvasSize.height - h) / 2),
					width: w,
					height: h,
					imageData: dataUrl,
				} as ImagePptxElement);
			};
			img.onerror = () => {
				addElement({
					id: generateElementId(),
					type: 'image',
					x: 100,
					y: 100,
					width: 400,
					height: 300,
					imageData: dataUrl,
				} as ImagePptxElement);
			};
			img.src = dataUrl;
		};
		reader.readAsDataURL(file);
	};

	/**
	 * Insert an audio or video file as a media element. Resolves once the element is added (after
	 * probing a video's size) or with the reason it was not; the toolbar's hidden input and a host's
	 * own picker both end here.
	 */
	const insertMediaFile = (
		file: File,
		shouldInsert?: () => boolean,
	): Promise<BuildingBlockResult> => {
		if (!activeSlide) {
			return Promise.resolve(buildingBlockFailure('no-slide'));
		}
		const mediaType = classifyMediaType(file.type);
		if (!mediaType) {
			return Promise.resolve(buildingBlockFailure('unsupported-type'));
		}
		const elFilePath: string | undefined = (file as File & { path?: string }).path;

		return new Promise((resolve) => {
			const insertMediaElement = (mediaSourceUrl: string, width: number, height: number) => {
				if (shouldInsert && !shouldInsert()) {
					resolve(buildingBlockFailure('stale-target'));
					return;
				}
				addElement({
					id: generateElementId(),
					type: 'media',
					mediaType,
					mediaMimeType: file.type || undefined,
					mediaData: mediaSourceUrl,
					x: Math.round((canvasSize.width - width) / 2),
					y: Math.round((canvasSize.height - height) / 2),
					width,
					height,
				} as MediaPptxElement);
				resolve(BUILDING_BLOCK_OK);
			};
			const probeVideoSize = (
				probeSrc: string,
				mediaSourceUrl: string,
				release: () => void,
			): void => {
				const probeVideo = document.createElement('video');
				probeVideo.preload = 'metadata';
				probeVideo.onloadedmetadata = () => {
					release();
					const mW = 640,
						mH = 360;
					let w = probeVideo.videoWidth || mW,
						h = probeVideo.videoHeight || mH;
					if (w > mW || h > mH) {
						const sc = Math.min(mW / w, mH / h);
						w = Math.round(w * sc);
						h = Math.round(h * sc);
					}
					insertMediaElement(mediaSourceUrl, w, h);
				};
				probeVideo.onerror = () => {
					release();
					insertMediaElement(mediaSourceUrl, 640, 360);
				};
				probeVideo.src = probeSrc;
			};

			if (elFilePath) {
				let normalized = elFilePath.replace(/\\/gu, '/');
				if (!normalized.startsWith('/')) {
					normalized = `/${normalized}`;
				}
				const mediaUrl = `pptx-resource://media${encodeURI(normalized)}`;
				if (mediaType === 'audio') {
					insertMediaElement(mediaUrl, 420, 64);
					return;
				}
				const probeUrl = URL.createObjectURL(file);
				probeVideoSize(probeUrl, mediaUrl, () => URL.revokeObjectURL(probeUrl));
				return;
			}

			const reader = new FileReader();
			reader.onerror = () => resolve(buildingBlockFailure('decode-failed'));
			reader.onload = () => {
				const dataUrl = reader.result as string;
				if (!dataUrl) {
					resolve(buildingBlockFailure('decode-failed'));
					return;
				}
				if (mediaType === 'audio') {
					insertMediaElement(dataUrl, 420, 64);
					return;
				}
				probeVideoSize(dataUrl, dataUrl, () => {});
			};
			reader.readAsDataURL(file);
		});
	};

	const handleMediaFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
		const file = e.target.files?.[0];
		if (!file || !activeSlide) {
			return;
		}
		e.target.value = '';
		void insertMediaFile(file);
	};

	return { handleImageFileChange, handleMediaFileChange, insertMediaFile };
}
