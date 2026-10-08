import { buildXml } from '../xml/index';
import type { VisioPackage } from './package';
import type { VisioPasteShapesEdit } from './edit-paste-commands';
import {
	clipboardResources,
	clipboardPageContext,
	assertClipboardResourceReferences,
} from './clipboard-resources';
import { clipboardXml, clipboardShapeRoot, assertClipboardXmlLimits } from './clipboard-xml';
import { duplicateVisioShapes } from './edit-duplicate';
import { fail } from './package-common';

/** Import detached source fragments after actual resource and page-context compatibility proof. */
export async function pasteVisioShapes(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioPasteShapesEdit,
	check: () => void,
): Promise<readonly string[]> {
	assertClipboardXmlLimits(edit.clipboard, check);
	const context = await clipboardPageContext(pkg, edit.pageId);
	if (
		context.scale !== edit.clipboard.sourceDrawingScale ||
		buildXml(clipboardXml(edit.clipboard.pageContext, check)) !== context.xml
	)
		fail(
			'CLIPBOARD_CONTEXT_MISMATCH',
			'Clipboard requires matching page settings, ordinal and drawing ratio.',
		);
	const actual = await clipboardResources(pkg);
	if (
		actual.length !== edit.clipboard.resources.length ||
		actual.some((resource, index) => {
			const captured = edit.clipboard.resources[index]!;
			return (
				captured.path !== resource.path ||
				buildXml(clipboardXml(captured.xml, check)) !== resource.xml
			);
		})
	)
		fail(
			'CLIPBOARD_RESOURCE_MISMATCH',
			'Clipboard font, style and theme definitions differ from target.',
		);
	const source = clipboardShapeRoot(edit.clipboard, check);
	assertClipboardResourceReferences(source, document);
	return duplicateVisioShapes(
		pkg,
		pagePaths,
		roots,
		document,
		{ ...edit, type: 'duplicate-shapes' },
		check,
		{ root: source, pageId: edit.clipboard.sourcePageId },
	);
}
