import { onOffElement, parseInteger, parseOnOff } from './simple-types';
import { first, getW, type XmlElement } from './xml';

export interface StyleMetadata {
	id: string;
	name?: string;
	basedOn?: string;
	isDefault?: boolean;
	customStyle?: boolean;
	uiPriority?: number;
	quickFormat?: boolean;
}

export function readStyleMetadata(element: XmlElement, id: string): StyleMetadata {
	const name = getW(first(element, 'name'), 'val');
	const basedOn = getW(first(element, 'basedOn'), 'val');
	const isDefault = parseOnOff(getW(element, 'default'));
	const customStyle = parseOnOff(getW(element, 'customStyle'));
	const uiPriority = parseInteger(getW(first(element, 'uiPriority'), 'val'));
	const quickFormat = onOffElement(first(element, 'qFormat'));
	return {
		id,
		...(name !== undefined && { name }),
		...(basedOn !== undefined && { basedOn }),
		...(isDefault !== undefined && { isDefault }),
		...(customStyle !== undefined && { customStyle }),
		...(uiPriority !== undefined && { uiPriority }),
		...(quickFormat !== undefined && { quickFormat }),
	};
}
