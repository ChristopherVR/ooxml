import { createOfficeGuid } from '../crypto/uuid';
import { parseXml } from '../xml/index';
import { clipboardId } from './clipboard-types';
import { captureVisioClipboard } from './clipboard';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { addMasterToDrawing } from './edit-master-add';
import { mapStencilStyles } from './edit-stencil-import';
import { fail } from './package-common';
import { VisioPackage } from './package';
import { related, visioXml } from './parts';
import { attribute, children } from './sheet';
import { ensureStencilIcon } from './stencil-icon';
import { stencilMasterEntry } from './stencil-master-part';
import type { StencilSession } from './stencil-edit-session';

export async function addSelectedStencilShape(
	source: Uint8Array,
	edit: { drawing: Uint8Array; pageId: string; shapeId: string; name: string },
	{ limits, check, deadline }: StencilSession,
): Promise<Uint8Array> {
	const { drawing, name } = edit;
	const pageId = clipboardId(edit.pageId);
	const shapeId = clipboardId(edit.shapeId);
	const snapshot = await captureVisioClipboard(drawing, pageId, [shapeId], { limits });
	const selected = snapshot.shapes[0];
	if (!selected || snapshot.masters)
		fail('UNSUPPORTED_EDIT_PACKAGE', 'Select a standalone shape or copy its original master.');
	const shape = parseXml(selected.xml).documentElement;
	const sheet = cells(shape);
	if (sheet.has('BeginX') || sheet.has('EndX') || sheet.has('LayerMember'))
		fail('UNSUPPORTED_EDIT_PACKAGE', 'Select a standalone 2D shape without a layer.');
	const width = numeric(sheet.get('Width'));
	const height = numeric(sheet.get('Height'));
	if (!(width > 0 && height > 0))
		fail('UNSUPPORTED_EDIT_PACKAGE', 'The selected shape needs a positive width and height.');
	// Normalize the position to local master-page coordinates and retain the shape's original ID,
	// so formulas that read another cell on the same shape can still resolve it.
	const locX = numeric(sheet.get('LocPinX'), width / 2);
	const locY = numeric(sheet.get('LocPinY'), height / 2);
	setCell(shape, 'PinX', locX, 'LocPinX');
	setCell(shape, 'PinY', locY, 'LocPinY');
	const pkg = await VisioPackage.open(drawing, limits);
	const sourceDocument = await visioXml(
		pkg,
		(await related(pkg, '', 'document'))!,
		'VisioDocument',
	);
	const uniqueId = createOfficeGuid();
	const result = await addMasterToDrawing(
		source,
		{
			uniqueId,
			name,
			entry(masters, id, uniqueName, relId) {
				const entry = stencilMasterEntry(
					masters,
					{
						master: { size: { width, height } },
						uniqueId,
						baseId: uniqueId,
					},
					id,
					uniqueName,
					relId,
				);
				ensureStencilIcon(entry);
			},
			contents(document, touch) {
				if (shape.namespaceURI !== document.namespaceURI)
					fail('UNSUPPORTED_EDIT_PACKAGE', 'The shape and stencil use different Visio namespaces.');
				const contents = parseXml(
					`<MasterContents xmlns="${document.namespaceURI}" xml:space="preserve"><Shapes/></MasterContents>`,
				).documentElement;
				mapStencilStyles(sourceDocument, document, touch)(shape);
				// Clipboard admission checks the source. Where the destination cannot represent a
				// font or palette index identically, refuse the copy instead of changing the look.
				for (const [section, cell] of [
					['Character', 'Font'],
					['Paragraph', 'BulletFont'],
				] as const)
					for (const rowSet of children(shape, 'Section').filter(
						(item) => attribute(item, 'N') === section,
					))
						for (const row of children(rowSet, 'Row')) {
							const font = children(row, 'Cell').find((item) => attribute(item, 'N') === cell);
							if (font && attribute(font, 'V') !== '0')
								fail('UNSUPPORTED_EDIT_PACKAGE', 'The selected shape uses a non-default font.');
						}
				contents.firstChild?.appendChild(contents.ownerDocument!.importNode(shape, true));
				return contents;
			},
		},
		limits,
		limits.maxInputBytes,
		deadline,
		check,
		true,
	);
	return result.bytes;
}
