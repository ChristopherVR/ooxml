import type { VisioStyle } from './model.js';
import { number, type Cells, type Report } from './sheet.js';
import { themeFormat, type ThemeResources } from './theme-resolve.js';

/**
 * MS-VSDX 2.4.4.170 names cached cap 0 Rounded, 1 Square, and 2 Extended:
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/b08a8361-c0c1-4d5f-8915-6810bcedfbc3
 * DrawingML cap tokens have an unambiguous normalized meaning. Office's omitted
 * cap default is flat (MS-OI29500 2.1.1212), despite the ECMA prose saying square:
 * https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5564035f-88ff-4ab8-bbfb-d2998a83324d
 */
export function lineCap(
	cells: Cells,
	resources: ThemeResources,
	report: Report,
): VisioStyle['lineCap'] {
	const cell = cells.get('LineCap');
	if (!cell) return undefined;
	if (cell.value === 'Themed') {
		const line = themeFormat(cells, 'Line', resources);
		if (!line) {
			report('unresolved-line-cap', 'The requested theme line cap could not be resolved.');
			return undefined;
		}
		const cap = line.getAttribute('cap') ?? 'flat';
		if (cap === 'rnd') return 'round';
		if (cap === 'flat') return 'butt';
		if (cap === 'sq') return 'square';
		report('unsupported-theme-line-cap', 'The selected theme line has an unsupported cap token.');
		return undefined;
	}
	if (cell.value !== undefined && cell.value.length > 128) {
		report('missing-cached-value', 'Cell LineCap exceeds the bounded numeric cache length.');
		return undefined;
	}
	const value = number(cells, 'LineCap', NaN, report);
	if (value === 0) return 'round';
	if (value === 1 || value === 2) {
		// Compatibility inference corroborated by libvisio 49fb9d3a9d21d4374cad782925e48c577a41f5be,
		// src/lib/VSDContentCollector.cpp:2992-3007. No implementation code is copied.
		// Microsoft's legacy names alone do not define these caps' exact endpoint geometry.
		report(
			'inferred-line-cap',
			'Legacy square/extended cap geometry uses an inferred compatibility mapping.',
		);
		return value === 1 ? 'butt' : 'square';
	}
	if (Number.isFinite(value))
		report('invalid-line-cap', 'The cached line cap is outside the supported enumeration.');
	return undefined;
}
