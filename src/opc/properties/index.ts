export * from './types.js';
export { CORE_PROPERTY_FIELDS, parseCoreProperties, writeCoreProperties } from './core.js';
export { parseAppProperties, writeAppProperties } from './app.js';
export {
	allocateCustomPropertyIds,
	parseCustomProperties,
	writeCustomProperties,
} from './custom.js';
