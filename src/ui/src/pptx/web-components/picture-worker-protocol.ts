import type { PictureRasterHeader, PptxBackgroundRemoval } from 'ooxml-core/pptx';
import type { PictureMarks } from './picture-segmentation';

import type { PictureCompression } from '../../form/picture-compression';
export type { PictureCompression } from '../../form/picture-compression';

export type PictureWorkerRequest =
	| { kind: 'open'; source: Blob; header: PictureRasterHeader }
	| { kind: 'compress'; options: PictureCompression }
	| { kind: 'remove'; marks: PictureMarks; preview: boolean };

export type PictureWorkerResult =
	| { kind: 'opened'; width: number; height: number }
	| { kind: 'compressed'; image: Blob; deletedCrop: boolean }
	| { kind: 'removed'; image: Blob; removal: PptxBackgroundRemoval }
	| { kind: 'error' };
