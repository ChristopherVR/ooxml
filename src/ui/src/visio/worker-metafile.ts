import type { VisioMetafileTreeConverter } from 'ooxml-core/visio';

/** Load the bundled metafile decoder only if a supported picture reaches the converter. */
export const convertVisioMetafile: VisioMetafileTreeConverter = async (...args) => {
	const { convertMetafileToSvgTree } = await import('emf-converter');
	return convertMetafileToSvgTree(...args);
};
