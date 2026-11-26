/** The Picture Format tab and the Format Picture crop fields. Keys are the English UI text. */
export const enPicture = {
	'tab.picture-format': 'Picture Format',
	Crop: 'Crop',
	'Crop (%)': 'Crop (%)',
	'The crop must leave part of the picture showing.':
		'The crop must leave part of the picture showing.',
	'Cropping is unavailable for legacy VML pictures.':
		'Cropping is unavailable for legacy VML pictures.',
	'This edit is unavailable for legacy VML pictures.':
		'This edit is unavailable for legacy VML pictures.',
} as const;
type Keys = keyof typeof enPicture;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(Object.keys(enPicture).map((key, index) => [key, values[index]!])) as Record<
		Keys,
		string
	>;
export const frPicture = fromValues([
	"Format de l'image",
	'Rogner',
	'Rognage (%)',
	"Le rognage doit laisser une partie de l'image visible.",
	'Le rognage n’est pas disponible pour les anciennes images VML.',
	'Cette modification n’est pas disponible pour les anciennes images VML.',
]);
export const dePicture = fromValues([
	'Bildformat',
	'Zuschneiden',
	'Zuschneiden (%)',
	'Beim Zuschneiden muss ein Teil des Bilds sichtbar bleiben.',
	'Ältere VML-Bilder können nicht zugeschnitten werden.',
	'Diese Bearbeitung ist für ältere VML-Bilder nicht verfügbar.',
]);
export const esPicture = fromValues([
	'Formato de imagen',
	'Recortar',
	'Recorte (%)',
	'El recorte debe dejar visible parte de la imagen.',
	'El recorte no está disponible para imágenes VML antiguas.',
	'Esta edición no está disponible para imágenes VML antiguas.',
]);
export const zhPicture = fromValues([
	'图片格式',
	'裁剪',
	'裁剪 (%)',
	'裁剪后必须保留部分图片可见。',
	'旧版 VML 图片不支持裁剪。',
	'旧版 VML 图片不支持此编辑操作。',
]);
