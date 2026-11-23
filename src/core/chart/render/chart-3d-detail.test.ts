import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { barShapeMesh } from './chart-3d-shapes';
import { chartSpaceKind, renderChartSpaceSvg } from './render-space';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const num = (values: number[]) =>
	`<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const cats = ['Mon', 'Tue', 'Wed', 'Thu'];
const cat = `<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
const ser = (idx: number, values: number[], color: string, extra = '') =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S${idx}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${extra}${cat}<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>${num(values)}</c:numCache></c:numRef></c:val></c:ser>`;
const axes = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`;

function renderXml(
	group: string,
	options: { view?: string; surfaces?: string; withAxes?: boolean } = {},
): string {
	const view = options.view ?? '<c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>';
	const surfaces = options.surfaces ?? '';
	const ax = options.withAxes === false ? '' : axes;
	const xml = `<c:chartSpace ${C} ${A}><c:chart>${view}${surfaces}<c:plotArea><c:layout/>${group}${ax}</c:plotArea></c:chart></c:chartSpace>`;
	return renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 400, height: 300 });
}

describe('3-D floor and wall formatting', () => {
	const group = `<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [4, 7, 3, 9], '4472C4')}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`;

	it('applies solid fill, noFill, and outline formatting to floor, backWall and sideWall', () => {
		const surfaces = [
			'<c:floor><c:spPr><a:solidFill><a:srgbClr val="D9E1F2"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:floor>',
			'<c:backWall><c:spPr><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:backWall>',
			'<c:sideWall><c:spPr><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr></c:sideWall>',
		].join('');
		const svg = renderXml(group, { surfaces });
		expect(svg).toMatch(
			/fill="#D9E1F2"[^>]*stroke="#1F4E79"[^>]*stroke-dasharray="[^"]+"[^>]*data-chart-part="floor"/,
		);
		expect(svg).toMatch(/fill="#FFF2CC"[^>]*stroke="none"[^>]*data-chart-part="backWall"/);
		expect(svg).toMatch(/fill="none"[^>]*stroke="#C00000"[^>]*data-chart-part="sideWall"/);
	});

	it('supports gradient fills on 3-D walls and floor', () => {
		const surfaces =
			'<c:backWall><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="BDD7EE"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></c:spPr></c:backWall>';
		const svg = renderXml(group, { surfaces });
		expect(svg).toContain('<linearGradient');
		expect(svg).toMatch(/fill="url\(#[^"]+\)"[^>]*data-chart-part="backWall"/);
	});
});

describe('3-D bar c:shape geometries', () => {
	const baseBounds = {
		v0: 0,
		v1: 50,
		c0: 10,
		c1: 30,
		z0: 5,
		z1: 25,
		horizontal: false,
		base: 0,
		extreme: 100,
	};
	const meta = { series: 0, point: 0, color: '#4472C4' };

	it('builds distinct meshes for cylinder, cone, coneToMax, pyramid and pyramidToMax', () => {
		const cyl = barShapeMesh('cylinder', baseBounds, meta)!;
		const cone = barShapeMesh('cone', baseBounds, meta)!;
		const coneMax = barShapeMesh('coneToMax', baseBounds, meta)!;
		const pyr = barShapeMesh('pyramid', baseBounds, meta)!;
		const pyrMax = barShapeMesh('pyramidToMax', baseBounds, meta)!;
		// Cylinder: 16 sides + 2 caps = 18 faces; cone: 16 triangular sides + 1 base cap = 17 faces.
		expect(cyl.faces).toHaveLength(18);
		expect(cone.faces).toHaveLength(17);
		expect(cone.faces.filter((f) => f.pts.length === 3)).toHaveLength(16);
		// coneToMax at half height is a frustum: 16 quad sides + 2 caps = 18 faces.
		expect(coneMax.faces).toHaveLength(18);
		// Pyramid: 4 triangular sides + 1 base cap = 5 faces; pyramidToMax frustum: 4 quads + 2 caps = 6.
		expect(pyr.faces).toHaveLength(5);
		expect(pyr.faces.filter((f) => f.pts.length === 3)).toHaveLength(4);
		expect(pyrMax.faces).toHaveLength(6);
		// Horizontal bar cone tapers along X.
		const hCone = barShapeMesh('cone', { ...baseBounds, horizontal: true }, meta)!;
		expect(hCone.faces).toHaveLength(17);
	});

	it('renders bar3DChart with group and series c:shape overrides', () => {
		for (const shape of ['cylinder', 'cone', 'coneToMax', 'pyramid', 'pyramidToMax']) {
			const svg = renderXml(
				`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [4, 10], '4472C4')}<c:shape val="${shape}"/><c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
			);
			expect(svg.match(/<g data-chart-series/g)).toHaveLength(2);
			expect(svg).not.toMatch(/NaN|Infinity/);
		}
		// Series-level shape override differs from default box.
		const boxSvg = renderXml(
			`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [4, 8], '4472C4')}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
		);
		const cylOverrideSvg = renderXml(
			`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [4, 8], '4472C4', '<c:shape val="cylinder"/>')}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
		);
		expect(cylOverrideSvg).not.toBe(boxSvg);
		// Cylinder has more polygons per bar than a 3-visible-face box.
		expect((cylOverrideSvg.match(/<polygon/g) ?? []).length).toBeGreaterThan(
			(boxSvg.match(/<polygon/g) ?? []).length,
		);
	});
});

describe('3-D pie slice explosion', () => {
	const pieView = '<c:view3D><c:rotX val="30"/></c:view3D>';

	it('offsets exploded slices along their bisector, exposes cut walls, and depth-sorts back to front', () => {
		const baseSvg = renderXml(
			`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [1, 1, 1, 1], '4472C4')}</c:pie3DChart>`,
			{ view: pieView, withAxes: false },
		);
		const explodedSvg = renderXml(
			`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [1, 1, 1, 1], '4472C4', '<c:explosion val="25"/>')}</c:pie3DChart>`,
			{ view: pieView, withAxes: false },
		);
		expect(explodedSvg).not.toBe(baseSvg);
		expect(explodedSvg).not.toMatch(/NaN|Infinity/);
		// Exploded pie draws inner radial cut walls in addition to outer rim walls.
		expect((explodedSvg.match(/<polygon/g) ?? []).length).toBeGreaterThan(
			(baseSvg.match(/<polygon/g) ?? []).length,
		);
		// Slice 0 (top-right, mid = -pi/4, back) is painted before slice 1 (bottom-right, mid = +pi/4, front).
		const top0 = explodedSvg.indexOf('data-chart-point="0"');
		const top1 = explodedSvg.indexOf('data-chart-point="1"');
		expect(top0).toBeGreaterThan(-1);
		expect(top1).toBeGreaterThan(top0);
	});

	it('supports point-level explosion overriding series explosion', () => {
		const baseSvg = renderXml(
			`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [4, 3, 2, 1], '4472C4')}</c:pie3DChart>`,
			{ view: pieView, withAxes: false },
		);
		const pointExpSvg = renderXml(
			`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [4, 3, 2, 1], '4472C4', '<c:dPt><c:idx val="2"/><c:explosion val="30"/></c:dPt>')}</c:pie3DChart>`,
			{ view: pieView, withAxes: false },
		);
		expect(pointExpSvg).not.toBe(baseSvg);
		expect(pointExpSvg.match(/<path d="M[^"]* A/g)).toHaveLength(4);
	});
});

describe('3-D surface wireframe and 2-D contour modes', () => {
	const surfSeries = [
		ser(0, [4, 7, 3, 9], 'FF0000'),
		ser(1, [2, 5, 6, 3], '0000FF'),
		ser(2, [1, 8, 2, 5], '00FF00'),
	].join('');

	it('draws surface3DChart wireframe with fill="none" and band-coloured strokes', () => {
		const svg = renderXml(
			`<c:surface3DChart><c:wireframe val="1"/>${surfSeries}<c:axId val="1"/><c:axId val="2"/></c:surface3DChart>`,
			{ view: '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="0"/></c:view3D>' },
		);
		expect(svg.match(/<g data-chart-series/g)).toHaveLength(6);
		expect(svg).toContain('fill="none" stroke="#');
	});

	it('draws surfaceChart as a 2-D contour grid in both filled and wireframe modes', () => {
		const filledSpace = parseChartSpace(
			`<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:layout/><c:surfaceChart>${surfSeries}<c:axId val="1"/><c:axId val="2"/></c:surfaceChart>${axes}</c:plotArea></c:chart></c:chartSpace>`,
		).chartSpace;
		expect(chartSpaceKind(filledSpace)).toEqual({ type: 'surface', drawn: true });
		const filled = renderChartSpaceSvg(filledSpace, { width: 400, height: 300 });
		expect(filled.match(/data-chart-part="contourCell"/g)).toHaveLength(6);
		expect(filled).toContain('>Mon<');
		expect(filled).toContain('>S0<');

		const wire = renderXml(
			`<c:surfaceChart><c:wireframe val="1"/>${surfSeries}<c:axId val="1"/><c:axId val="2"/></c:surfaceChart>`,
			{ view: '' },
		);
		expect(
			wire.match(
				/fill="none" stroke="#[0-9A-F]{6}" stroke-width="1" data-chart-part="contourCell"/gi,
			),
		).toHaveLength(6);
	});
});

describe('3-D chart data labels', () => {
	const dLblsVal = '<c:dLbls><c:showVal val="1"/></c:dLbls>';

	it('draws projected data labels on bar3D (vertical and horizontal), line3D and area3D', () => {
		const colSvg = renderXml(
			`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [14, 27, 33, 49], '4472C4')}${dLblsVal}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
		);
		for (const v of ['14', '27', '33', '49']) expect(colSvg).toContain(`>${v}</text>`);

		const hBarSvg = renderXml(
			`<c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/>${ser(0, [11, 22], '4472C4')}<c:dLbls><c:dLblPos val="ctr"/><c:showVal val="1"/></c:dLbls><c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
		);
		expect(hBarSvg).toContain('>11</text>');
		expect(hBarSvg).toContain('>22</text>');

		const lineSvg = renderXml(
			`<c:line3DChart><c:grouping val="standard"/>${ser(0, [15, 25, 35, 45], 'ED7D31')}${dLblsVal}<c:axId val="1"/><c:axId val="2"/></c:line3DChart>`,
			{ view: '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="0"/></c:view3D>' },
		);
		for (const v of ['15', '25', '35', '45']) expect(lineSvg).toContain(`>${v}</text>`);

		const areaSvg = renderXml(
			`<c:area3DChart><c:grouping val="standard"/>${ser(0, [12, 24, 36, 48], '70AD47')}${dLblsVal}<c:axId val="1"/><c:axId val="2"/></c:area3DChart>`,
		);
		for (const v of ['12', '24', '36', '48']) expect(areaSvg).toContain(`>${v}</text>`);
	});

	it('draws data labels and leader lines on pie3DChart', () => {
		const svg = renderXml(
			`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [40, 30, 20, 10], '4472C4')}<c:dLbls><c:dLblPos val="outEnd"/><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls></c:pie3DChart>`,
			{ view: '<c:view3D><c:rotX val="30"/></c:view3D>', withAxes: false },
		);
		for (const pct of ['40%', '30%', '20%', '10%']) expect(svg).toContain(`>${pct}</text>`);
		expect(svg.match(/data-chart-part="leaderLine"/g)).toHaveLength(4);
	});
});
