import { sha256 } from '../digest/sha256';

/** Marks a generated approximation so consumers never describe it as a native Office cache. */
export const GENERATED_DIAGRAM_CACHE_NS = 'urn:ooxml-core:diagram:generated-cache';
export const GENERATED_DIAGRAM_CACHE_NOTICE =
	'Generated SmartArt preview; not native Office layout.';
export const generatedDrawingDigest = (xml: string): string =>
	Array.from(sha256(new TextEncoder().encode(xml)), (byte) =>
		byte.toString(16).padStart(2, '0'),
	).join('');

export const DEFAULT_GENERATED_DIAGRAM_SCHEME: Readonly<Record<string, string>> = {
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	tx1: '#000000',
	bg1: '#FFFFFF',
	tx2: '#44546A',
	bg2: '#E7E6E6',
};
