import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { createChart3DCamera } from './chart-3d-camera';
import { boxMesh, extrudeMesh } from './chart-3d-mesh';
import { chart3DPieGeometry } from './chart-3d-pie';
import { resolveChart3DProjection } from './chart-3d-projection';
import { chart3DFaceVisible, chart3DMeshesSvg } from './chart-3d-scene';
import { buildChart3DBoxFaceColors, shadeChart3DFace } from './chart-3d-shading';
import { chartSpaceKind, renderChartSpaceSvg } from './render-space';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const num = (values: number[]) =>
	`<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const cats = ['Mon', 'Tue', 'Wed', 'Thu'];
const cat = `<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
const ser = (idx: number, values: number[], color: string) =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S${idx}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${cat}<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>${num(values)}</c:numCache></c:numRef></c:val></c:ser>`;
const axes = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`;
const series = [ser(0, [4, 7, 3, 9], 'FF0000'), ser(1, [2, 5, 6, 3], '0000FF')].join('');

function chart(group: string, view = '<c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>') {
	const xml = `<c:chartSpace ${C} ${A}><c:chart>${view}<c:plotArea><c:layout/>${group}${axes}</c:plotArea></c:chart></c:chartSpace>`;
	return parseChartSpace(xml).chartSpace;
}
const render = (space: ReturnType<typeof chart>) =>
	renderChartSpaceSvg(space, { width: 400, height: 300 });
const bar3D = (grouping: string, dir = 'col', extra = '') =>
	chart(
		`<c:bar3DChart><c:barDir val="${dir}"/><c:grouping val="${grouping}"/>${series}${extra}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
	);

describe('3-D projection', () => {
	it('uses an oblique projection for bar3D and a perspective camera elsewhere', () => {
		expect(resolveChart3DProjection('bar3D', undefined).mode).toBe('oblique');
		expect(resolveChart3DProjection('line3D', undefined).mode).toBe('perspective');
		expect(resolveChart3DProjection('bar3D', { rAngAx: false }).mode).toBe('perspective');
	});

	it('defaults the tilt per type and derives the oblique shear', () => {
		expect(resolveChart3DProjection('pie3D', undefined)).toMatchObject({ rotXDeg: 30, rotYDeg: 0 });
		const p = resolveChart3DProjection('bar3D', { rotX: 30, rotY: 30 });
		expect(p.mode).toBe('oblique');
		if (p.mode === 'oblique') expect(p.shearX).toBeCloseTo(0.5);
		if (p.mode === 'oblique') expect(p.shearY).toBeCloseTo(-0.5);
	});

	it('widens the field of view with the perspective angle', () => {
		const near = resolveChart3DProjection('line3D', { perspective: 0 });
		const far = resolveChart3DProjection('line3D', { perspective: 120 });
		if (near.mode === 'perspective' && far.mode === 'perspective') {
			expect(near.fovDeg).toBe(15);
			expect(far.fovDeg).toBe(75);
		}
	});
});

describe('3-D shading', () => {
	it('keeps the front colour and darkens tops and sides', () => {
		expect(shadeChart3DFace('#FF0000', 'front')).toBe('#FF0000');
		expect(shadeChart3DFace('#FF0000', 'top')).not.toBe('#FF0000');
		expect(shadeChart3DFace('url(#g)', 'right')).toBe('url(#g)');
		expect(buildChart3DBoxFaceColors('#FF0000')).toHaveLength(6);
	});
});

describe('3-D camera and meshes', () => {
	const fit = { x: 0, y: 0, w: 200, h: 100 };
	const size = { w: 100, h: 50, d: 40 };

	it('fits the projected box inside the rectangle for both projections', () => {
		for (const mode of ['bar3D', 'line3D']) {
			const cam = createChart3DCamera(resolveChart3DProjection(mode, undefined), size, fit);
			for (const x of [0, 100])
				for (const y of [0, 50])
					for (const z of [0, 40]) {
						const q = cam.project([x, y, z]);
						expect(q.x).toBeGreaterThanOrEqual(-1e-6);
						expect(q.x).toBeLessThanOrEqual(200 + 1e-6);
						expect(q.y).toBeGreaterThanOrEqual(-1e-6);
						expect(q.y).toBeLessThanOrEqual(100 + 1e-6);
					}
		}
	});

	it('draws the back of the box higher on screen when seen from above', () => {
		const cam = createChart3DCamera(resolveChart3DProjection('bar3D', undefined), size, fit);
		expect(cam.project([0, 0, 40]).y).toBeLessThan(cam.project([0, 0, 0]).y);
	});

	it('culls faces turned away: three faces of a box are visible from above-right', () => {
		for (const mode of ['bar3D', 'line3D']) {
			const cam = createChart3DCamera(resolveChart3DProjection(mode, undefined), size, fit);
			const box = boxMesh(10, 30, 0, 20, 10, 20, { series: 0, point: 0, color: '#FF0000' })!;
			const visible = box.faces.filter((f) => chart3DFaceVisible(cam, f)).map((f) => f.kind);
			expect(visible.sort()).toEqual(['front', 'right', 'top']);
		}
	});

	it('draws far marks before near ones and skips degenerate outlines', () => {
		const cam = createChart3DCamera(resolveChart3DProjection('bar3D', undefined), size, fit);
		const far = boxMesh(0, 10, 0, 10, 30, 40, { series: 1, point: 0, color: '#00FF00' })!;
		const near = boxMesh(0, 10, 0, 10, 0, 10, { series: 0, point: 0, color: '#FF0000' })!;
		const svg = chart3DMeshesSvg(cam, [near, far]);
		expect(svg.indexOf('data-chart-series="1"')).toBeLessThan(svg.indexOf('data-chart-series="0"'));
		expect(boxMesh(0, 10, 5, 5, 0, 10, { series: 0, point: 0, color: '#000000' })).toBeUndefined();
		expect(extrudeMesh([[0, 0]], 0, 1, { series: 0, point: 0, color: '#000000' })).toBeUndefined();
	});

	it('gives the pie a thickness that vanishes when seen from above', () => {
		const area = { x: 0, y: 0, w: 200, h: 200 };
		const tilted = chart3DPieGeometry(area, { rotX: 30 });
		const top = chart3DPieGeometry(area, { rotX: 90 });
		expect(tilted.ry).toBeLessThan(tilted.rx);
		expect(tilted.thickness).toBeGreaterThan(0);
		expect(top.thickness).toBeCloseTo(0);
		expect(top.ry).toBeCloseTo(top.rx);
	});
});

describe('3-D chart SVG', () => {
	it('draws a 3-D clustered column with shaded side and top faces', () => {
		const svg = render(bar3D('clustered'));
		expect(svg).toContain('data-chart-series="0"');
		// 8 bars, each with a front, top and one side polygon, plus the walls.
		const bars = svg.match(/<g data-chart-series/g) ?? [];
		expect(bars).toHaveLength(8);
		expect(svg).toContain('fill="#FF0000"');
		expect(svg).toContain('fill="#bf0000"');
		expect(svg).toContain('fill="#a30000"');
		expect(svg).toContain('Mon');
	});

	it('stacks 3-D columns and puts standard series in depth rows', () => {
		const stacked = render(bar3D('stacked'));
		expect(stacked.match(/<g data-chart-series/g)).toHaveLength(8);
		const standard = render(bar3D('standard'));
		expect(standard.match(/<g data-chart-series/g)).toHaveLength(8);
		expect(standard).not.toBe(stacked);
	});

	it('draws a 3-D bar chart horizontally', () => {
		const svg = render(bar3D('clustered', 'bar'));
		expect(svg.match(/<g data-chart-series/g)).toHaveLength(8);
		expect(svg).toContain('>Mon<');
	});

	it('draws line3DChart as ribbons per segment', () => {
		const svg = render(
			chart(
				`<c:line3DChart><c:grouping val="standard"/>${series}<c:axId val="1"/><c:axId val="2"/></c:line3DChart>`,
				'<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="0"/></c:view3D>',
			),
		);
		expect(svg.match(/<g data-chart-series/g)).toHaveLength(6);
	});

	it('draws area3DChart as one prism per series', () => {
		const svg = render(
			chart(
				`<c:area3DChart><c:grouping val="standard"/>${series}<c:axId val="1"/><c:axId val="2"/></c:area3DChart>`,
			),
		);
		expect(svg.match(/<g data-chart-series/g)).toHaveLength(2);
		expect(svg).toContain('fill="#FF0000"');
	});

	it('draws pie3DChart as walls under tilted tops', () => {
		const svg = render(
			chart(
				`<c:pie3DChart><c:varyColors val="1"/>${ser(0, [4, 3, 2, 1], '4472C4')}</c:pie3DChart>`,
				'<c:view3D><c:rotX val="30"/></c:view3D>',
			),
		);
		expect(svg).toContain('<polygon');
		expect(svg.match(/<path d="M[^"]* A/g)?.length).toBe(4);
		expect(svg.indexOf('<polygon')).toBeLessThan(svg.indexOf('<path d="M'));
	});

	it('draws surface3DChart as banded patches and reports it drawn', () => {
		const space = chart(
			`<c:surface3DChart>${series}${ser(2, [1, 8, 2, 5], '00FF00')}<c:axId val="1"/><c:axId val="2"/></c:surface3DChart>`,
			'<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="0"/></c:view3D>',
		);
		const svg = render(space);
		expect(svg.match(/<g data-chart-series/g)?.length).toBe(6);
		expect(chartSpaceKind(space)).toMatchObject({ type: 'surface', drawn: true });
	});

	it('still draws flat charts and an unviewed surface the old way', () => {
		const flat = render(
			chart(
				`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series}<c:axId val="1"/><c:axId val="2"/></c:barChart>`,
				'',
			),
		);
		expect(flat).toContain('<rect');
		expect(flat).not.toContain('<polygon');
		const surface = chart(
			`<c:surfaceChart>${series}<c:axId val="1"/><c:axId val="2"/></c:surfaceChart>`,
			'',
		);
		expect(render(surface)).toContain('surface charts are not drawn');
	});
});
