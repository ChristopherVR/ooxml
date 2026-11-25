import type { XmlObject } from '../../types';
import { attrByLocalName, childByLocalName, childrenByLocalName, localName } from './image-a14-xml';
import { SVG_BLIP_EXT_URI } from './picture-svg-blip';

export function pictureBlipFill(shape: XmlObject | undefined): XmlObject | undefined {
	return (
		childByLocalName(shape, 'blipFill') ??
		childByLocalName(childByLocalName(shape, 'spPr'), 'blipFill')
	);
}

/** Reuse the original picture tree, including non-visual properties and fill layout. */
export function rebindPictureBlip(shape: XmlObject, relationshipId: string): XmlObject | undefined {
	const edited = structuredClone(shape);
	const blip = childByLocalName(pictureBlipFill(edited), 'blip');
	if (!blip) return undefined;
	for (const key of Object.keys(blip)) {
		if (key.startsWith('@_') && ['embed', 'link'].includes(localName(key.slice(2))))
			delete blip[key];
	}
	blip['@_r:embed'] = relationshipId;
	const extList = childByLocalName(blip, 'extLst');
	if (extList) {
		const extensions = childrenByLocalName(extList, 'ext');
		const remaining = extensions.filter((ext) => attrByLocalName(ext, 'uri') !== SVG_BLIP_EXT_URI);
		if (remaining.length !== extensions.length) {
			for (const key of Object.keys(extList)) {
				if (localName(key) === 'ext') delete extList[key];
			}
			if (remaining.length) extList['a:ext'] = remaining;
		}
	}
	return edited;
}
