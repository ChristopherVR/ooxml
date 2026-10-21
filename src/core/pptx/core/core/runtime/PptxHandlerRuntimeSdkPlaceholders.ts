import type { PptxElement, XmlObject } from '../../types';
import { placeholderContentRole } from '../../utils/placeholder-remap';
import { xmlPath } from '../../utils/xml-access';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeBackgroundStyles';
import { applyPlaceholderFrame, placeholderShape, selectTextPlaceholder } from './sdk-placeholder';
import type { SdkPlaceholderOptions } from './sdk-placeholder';

/** Layout-bound SDK content, resolved without creating slide parts or relationships. */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	public createPlaceholderElement(
		layoutPath: string,
		input: PptxElement,
		type: string,
		options: SdkPlaceholderOptions,
	): PptxElement {
		if (input.type !== 'text') throw new Error('Placeholder input must be text');
		const layout = this.layoutXmlMap.get(layoutPath);
		if (!layout) throw new Error(`Unknown layout: ${layoutPath}`);
		const identity = selectTextPlaceholder(
			input,
			type,
			options,
			this.extractLayoutPlaceholders(layout, layoutPath),
		);
		const shape = placeholderShape(
			identity,
			this.createParagraphsFromTextContent(input.text, undefined, input.textSegments),
		);
		// The optional layout context uses the loader's unchanged lookup rules.
		// No synthetic slide path or temporary relationship-map entry is needed.
		const inherited = this.findPlaceholderNode(
			'',
			{ type: identity.type.toLowerCase(), idx: identity.idx },
			layoutPath,
		);
		const fallback = this.sdkMasterFrame(layoutPath, identity.type, identity.idx, inherited);
		// Some templates use an obj layout slot backed by a body master slot
		// with a different index. Keep that fallback local to newly authored
		// content and write its frame explicitly: the existing loader must not
		// change how it interprets any pre-existing slide.
		applyPlaceholderFrame(
			shape,
			fallback ?? inherited,
			options,
			PptxHandlerRuntime.EMU_PER_PX,
			fallback !== undefined,
		);
		const element = this.parseShape(shape, input.id, undefined, layoutPath);
		if (!element) throw new Error('Unable to resolve placeholder content');
		// The parser leaves textStyle unset for empty text, so save would rebuild
		// the run without the inheritance baseline and stamp a default language
		// (dropping authored properties such as bold). The first segment already
		// carries the resolved style and baseline; use it for empty placeholders.
		if (
			(element.type === 'text' || element.type === 'shape') &&
			!element.textStyle &&
			element.textSegments?.[0]?.style
		) {
			element.textStyle = { ...element.textSegments[0].style };
		}
		return element;
	}

	private sdkMasterFrame(
		layoutPath: string,
		type: string,
		idx: string | undefined,
		inherited: XmlObject | undefined,
	): XmlObject | undefined {
		if (!inherited) return undefined;
		if (
			xmlPath(inherited, 'p:spPr', 'a:xfrm', 'a:off') &&
			xmlPath(inherited, 'p:spPr', 'a:xfrm', 'a:ext')
		)
			return undefined;
		const masterPath = this.resolveMasterPathForLayout(layoutPath);
		const tree = masterPath
			? xmlPath(this.masterXmlMap.get(masterPath), 'p:sldMaster', 'p:cSld', 'p:spTree')
			: undefined;
		const role = placeholderContentRole(type.toLowerCase());
		const candidates = (this.ensureArray(tree?.['p:sp']) as XmlObject[]).filter((shape) => {
			const info = this.extractPlaceholderInfo(xmlPath(shape, 'p:nvSpPr', 'p:nvPr'));
			return info && placeholderContentRole(info.type) === role;
		});
		const master =
			candidates.find(
				(shape) => this.extractPlaceholderInfo(xmlPath(shape, 'p:nvSpPr', 'p:nvPr'))?.idx === idx,
			) ?? candidates[0];
		return master ? this.mergeXmlObjects(master, inherited) : undefined;
	}
}
