import { isSelfContainedGlb } from '../../drawingml/index';
import { MAX_LOCAL_OBJECT_BYTES, type EmbeddableOfficePackage } from '../../opc/index';
import type { DrawingAnchor, EmbeddedDocumentObject, Model3DObject } from '../model';
import { newMediaPart } from './charts';
import { sheetAt, type EditContext } from './context';
import { checkAnchor } from './drawings';

interface LocalObjectInputs {
	name: string;
	bytes: Uint8Array;
	preview: Uint8Array;
	anchor: DrawingAnchor;
}

export type LocalObjectInsertion = LocalObjectInputs &
	({ kind: 'model3d' } | { kind: 'object'; format: EmbeddableOfficePackage });

/** Inserts a fully local GLB or an embedded Office document as one undoable drawing edit. */
export function insertLocalObject(
	ctx: EditContext,
	s: number,
	input: LocalObjectInsertion,
): number {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	checkAnchor(input.anchor);
	if (input.bytes.length > MAX_LOCAL_OBJECT_BYTES)
		throw new RangeError('The selected file is too large.');
	const preview = input.preview.slice();
	const bytes = input.bytes.slice();
	const anchor = structuredClone(input.anchor);
	const format = input.kind === 'object' ? input.format : undefined;
	if (!format && !isSelfContainedGlb(bytes))
		throw new RangeError('Select a self-contained GLB model.');
	const kind = format ? 'embeddedObject' : 'model3d';
	const folder = format ? 'embeddings' : 'media';
	const base = format ? 'oleObject' : 'model3d';
	const extension = format?.extension ?? 'glb';
	const taken = new Set(Array.from(workbook.source?.parts.keys() ?? [], (p) => p.toLowerCase()));
	let n = 1;
	while (taken.has(`xl/${folder}/${base}${n}.${extension}`)) n++;
	const partName = `xl/${folder}/${base}${n}.${extension}`;
	const previewPartName = newMediaPart(workbook, 'image/png');
	const common = {
		name: input.name,
		anchor,
		partName,
		previewPartName,
		previewContentType: 'image/png',
	};
	const object: Model3DObject | EmbeddedDocumentObject = format
		? { ...common, kind: 'embeddedObject', progId: format.progId, contentType: format.mime }
		: { ...common, kind: 'model3d' };
	return ctx.run(
		kind === 'model3d' ? 'Insert 3D model' : 'Insert Object',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			workbook.source ??= { parts: new Map() };
			workbook.source.parts.set(partName, bytes);
			workbook.source.parts.set(previewPartName, preview);
			sheet.drawings.push(object);
			return sheet.drawings.length - 1;
		},
		{ sheet: s },
	);
}
