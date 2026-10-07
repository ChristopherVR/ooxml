type ConvertMetafileToDataUrl = typeof import('emf-converter').convertMetafileToDataUrl;

/**
 * Converts an EMF or WMF picture to a data URL, loading `emf-converter` on first use.
 *
 * Most decks have no metafile pictures, and the converter is one of the largest modules
 * the pptx area reaches, so a static import would make every consumer download it.
 */
export const convertMetafileToDataUrl: ConvertMetafileToDataUrl = async (...args) => {
	const { convertMetafileToDataUrl: convert } = await import('emf-converter');
	return convert(...args);
};
