import { PresentationBuilder } from './PresentationBuilder';

export interface PhotoAlbumPicture {
	/** A local embedded image data URL. */
	data: string;
	width: number;
	height: number;
	caption: string;
}

export interface PhotoAlbumOptions {
	title: string;
	pictures: readonly PhotoAlbumPicture[];
	picturesPerSlide: 1 | 2 | 4;
	captions?: boolean;
}

/** Build an independent deck containing editable pictures and optional text captions. */
export async function createPhotoAlbum(options: PhotoAlbumOptions): Promise<Uint8Array> {
	const { title, pictures, picturesPerSlide, captions } = options;
	if (
		pictures.length === 0 ||
		![1, 2, 4].includes(picturesPerSlide) ||
		pictures.some(
			({ data, width, height }) =>
				!data.startsWith('data:image/') ||
				!Number.isFinite(width) ||
				!Number.isFinite(height) ||
				width <= 0 ||
				height <= 0,
		)
	)
		throw new Error('Invalid photo album pictures');
	const { handler, data, createSlide } = await PresentationBuilder.create({ title });
	try {
		const { width, height } = data;
		const margin = Math.round(Math.min(width, height) * 0.06);
		const columns = picturesPerSlide === 1 ? 1 : 2;
		const rows = picturesPerSlide / columns;
		const slotWidth = (width - (columns + 1) * margin) / columns;
		const slotHeight = (height - (rows + 1) * margin) / rows;
		const captionHeight = captions ? Math.min(36, slotHeight / 5) : 0;
		data.slides.push(
			createSlide('Blank')
				.addText(title, {
					x: margin,
					y: height / 3,
					width: width - 2 * margin,
					height: height / 3,
					fontSize: 36,
					bold: true,
					alignment: 'center',
				})
				.build(),
		);
		for (let start = 0; start < pictures.length; start += picturesPerSlide) {
			const slide = createSlide('Blank');
			for (const [index, picture] of pictures.slice(start, start + picturesPerSlide).entries()) {
				const left = margin + (index % columns) * (slotWidth + margin);
				const top = margin + Math.floor(index / columns) * (slotHeight + margin);
				const pictureArea = slotHeight - captionHeight;
				const scale = Math.min(slotWidth / picture.width, pictureArea / picture.height);
				const imageWidth = picture.width * scale;
				const imageHeight = picture.height * scale;
				slide.addImage(picture.data, {
					x: left + (slotWidth - imageWidth) / 2,
					y: top + (pictureArea - imageHeight) / 2,
					width: imageWidth,
					height: imageHeight,
					altText: picture.caption,
				});
				if (captions) {
					slide.addText(picture.caption, {
						x: left,
						y: top + pictureArea,
						width: slotWidth,
						height: captionHeight,
						fontSize: 14,
						alignment: 'center',
					});
				}
			}
			data.slides.push(slide.build());
		}
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}
