export const enPictureWatermark = {
	'Picture watermark': 'Picture watermark',
	'Select Picture…': 'Select Picture…',
	Scale: 'Scale',
	Auto: 'Auto',
	Washout: 'Washout',
	'Select a picture for the watermark.': 'Select a picture for the watermark.',
} as const;
type Keys = keyof typeof enPictureWatermark;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enPictureWatermark).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frPictureWatermark = fromValues([
	'Filigrane image',
	'Sélectionner l’image…',
	'Échelle',
	'Automatique',
	'Estompé',
	'Sélectionnez une image pour le filigrane.',
]);
export const dePictureWatermark = fromValues([
	'Bildwasserzeichen',
	'Bild auswählen…',
	'Skalieren',
	'Automatisch',
	'Auswaschen',
	'Wählen Sie ein Bild für das Wasserzeichen aus.',
]);
export const esPictureWatermark = fromValues([
	'Marca de agua de imagen',
	'Seleccionar imagen…',
	'Escala',
	'Automática',
	'Decolorar',
	'Seleccione una imagen para la marca de agua.',
]);
export const zhPictureWatermark = fromValues([
	'图片水印',
	'选择图片…',
	'缩放',
	'自动',
	'冲蚀',
	'请选择水印图片。',
]);
