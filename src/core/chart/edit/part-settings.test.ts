import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import {
	setChartColorMapOverride,
	setChartExternalDataAutoUpdate,
	setChartPivotFormatsXml,
	setChartPrintSettingsXml,
	setChartUserShapesReference,
} from './part-settings';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const load = (head: string, tail = '') =>
	parseChartSpace(
		`<c:chartSpace xmlns:c="${C}" xmlns:r="${R}">${head}<c:chart><c:plotArea/></c:chart>${tail}</c:chartSpace>`,
	).chartSpace;
const roles = (value: string) =>
	Object.fromEntries(
		[
			'bg1',
			'tx1',
			'bg2',
			'tx2',
			'accent1',
			'accent2',
			'accent3',
			'accent4',
			'accent5',
			'accent6',
			'hlink',
			'folHlink',
		].map((role) => [role, role.startsWith('accent') || role.endsWith('link') ? role : value]),
	);

describe('chart part settings', () => {
	it('models the colour mapping and pivot formats as written', () => {
		const { chartSpace, issues } = parseChartSpace(
			`<c:chartSpace xmlns:c="${C}"><c:clrMapOvr bg1="lt1"/><c:chart><c:autoTitleDeleted val="0"/><c:pivotFmts/><c:plotArea/></c:chart></c:chartSpace>`,
		);
		expect(issues).toEqual([]);
		const xml = writeChartSpace(chartSpace);
		expect(xml).toMatch(
			/<c:clrMapOvr bg1="lt1"\/><c:chart><c:autoTitleDeleted val="0"\/><c:pivotFmts\/>/u,
		);
	});

	it('sets, replaces and removes a colour mapping', () => {
		const set = setChartColorMapOverride(load(''), roles('dk1'));
		expect(writeChartSpace(set)).toMatch(/<c:clrMapOvr bg1="dk1" tx1="dk1"[^>]*\/><c:chart>/u);
		expect(writeChartSpace(setChartColorMapOverride(set, null))).not.toContain('clrMapOvr');
		expect(() => setChartColorMapOverride(load(''), { bg1: 'dk1' })).toThrow(/12 roles/u);
		expect(() => setChartColorMapOverride(load(''), roles('red'))).toThrow(RangeError);
	});

	it('sets the auto-update flag of an external data reference only', () => {
		const space = load('', '<c:externalData r:id="rId1"/>');
		expect(writeChartSpace(setChartExternalDataAutoUpdate(space, true))).toContain(
			'<c:externalData r:id="rId1"><c:autoUpdate val="1"/></c:externalData>',
		);
		expect(() => setChartExternalDataAutoUpdate(load(''), true)).toThrow(/external data/u);
	});

	it('places encoded print settings, pivot formats and the user shapes reference', () => {
		let space = setChartPrintSettingsXml(
			load(''),
			`<c:printSettings xmlns:c="${C}"><c:pageSetup orientation="portrait"/></c:printSettings>`,
		);
		space = setChartPivotFormatsXml(space, `<c:pivotFmts xmlns:c="${C}"/>`);
		space = setChartUserShapesReference(space, 'rId4');
		const xml = writeChartSpace(space);
		expect(xml).toMatch(/<c:pivotFmts\/><c:plotArea/u);
		expect(xml).toMatch(/<c:printSettings>.*<\/c:printSettings><c:userShapes r:id="rId4"\/>/u);
		expect(() => setChartPrintSettingsXml(space, `<c:pivotFmts xmlns:c="${C}"/>`)).toThrow(
			RangeError,
		);
		expect(writeChartSpace(setChartUserShapesReference(space, null))).not.toContain('userShapes');
	});
});
