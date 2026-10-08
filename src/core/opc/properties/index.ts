export * from './types';
export {
	CORE_PROPERTY_FIELDS,
	formatW3cdtf,
	parseCoreProperties,
	writeCoreProperties,
} from './core';
export { STRICT_APP_NAMESPACES, parseAppProperties, writeAppProperties } from './app';
export {
	allocateCustomPropertyIds,
	customPropertyFromText,
	parseCustomProperties,
	parseCustomPropertyTexts,
	writeCustomProperties,
} from './custom';
export type { CustomPropertyText } from './custom';
export { fromTitleGroups, replaceTitleGroup, titleGroups } from './title-groups';
export type { TitleGroup } from './title-groups';
