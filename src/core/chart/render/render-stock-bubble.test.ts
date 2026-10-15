import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { renderChartSpaceSvg } from './render-space';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const num = (values: number[]) =>
	`<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}`;
const label = (text: string) =>
	`<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
const cats = `<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>${['Mon', 'Tue', 'Wed', 'Thu'].map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
const ser = (i: number, name: string, values: number[]) =>
	`<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${label(name)}${cats}<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>${num(values)}</c:numCache></c:numRef></c:val></c:ser>`;
const axes = (id: number, cross: number, position: string, extra = '') =>
	`<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${position === 'r' ? 1 : 0}"/><c:axPos val="b"/><c:crossAx val="${cross}"/></c:catAx>
<c:valAx><c:axId val="${cross}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/><c:crossAx val="${id}"/>${extra}</c:valAx>`;

const ohlc = [
	ser(0, 'Open', [10, 11, 12, 11]),
	ser(1, 'High', [14, 15, 15, 14]),
	ser(2, 'Low', [9, 10, 10, 9]),
	ser(3, 'Close', [12, 13, 11, 13]),
].join('');

const render = (xml: string) =>
	renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 480, height: 300 });

describe('stock chart formatting', () => {
	const styled = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:layout/><c:stockChart>${ohlc}
<c:hiLowLines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:hiLowLines>
<c:upDownBars><c:gapWidth val="100"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:upBars>
<c:downBars><c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:downBars></c:upDownBars>
<c:axId val="1"/><c:axId val="2"/></c:stockChart>${axes(1, 2, 'l')}</c:plotArea></c:chart></c:chartSpace>`;

	it("paints the chart's own wick, bar fills, outlines and gap width", () => {
		const svg = render(styled);
		const candles = svg.match(/<g data-chart-series="3"[\s\S]*?<\/g>/g)!;
		expect(candles).toHaveLength(4);
		// Wick: red, 2pt wide (19050 EMU is 2px).
		expect(candles[0]).toContain('stroke="#FF0000" stroke-width="2"');
		// Up bars: green without an outline; down bars: red with a black 2px outline.
		expect(candles[0]).toMatch(/<rect [^>]*fill="#00AA00"(?![^>]*stroke=)/);
		expect(candles[2]).toMatch(/<rect [^>]*fill="#AA0000"[^>]*stroke="#000000" stroke-width="2"/);
		// A gap width of 100 fills half of the category band.
		const widths = candles.map((g) => Number(/<rect [^>]*\bwidth="([\d.]+)"/.exec(g)![1]));
		const wide = render(styled.replace('<c:gapWidth val="100"/>', '<c:gapWidth val="300"/>'));
		const narrow = Number(
			/<rect [^>]*\bwidth="([\d.]+)"/.exec(
				wide.match(/<g data-chart-series="3"[\s\S]*?<\/g>/)![0],
			)![1],
		);
		expect(narrow).toBeCloseTo(widths[0]! / 2, 1);
	});

	it('falls back to Excel defaults when the chart writes no formatting', () => {
		const plain = styled.replace(
			/<c:hiLowLines>[\s\S]*<\/c:upDownBars>/,
			'<c:hiLowLines/><c:upDownBars/>',
		);
		const candles = render(plain).match(/<g data-chart-series="3"[\s\S]*?<\/g>/g)!;
		expect(candles[0]).toContain('stroke="#334155"');
		expect(candles[0]).toContain('fill="#FFFFFF"');
		expect(candles[2]).toContain('fill="#404040"');
	});
});

describe('volume stock chart', () => {
	const volume = `<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'Volume', [1000, 2000, 1500, 1200])}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:stockChart>${ohlc.replace(/<c:idx val="(\d)"\/><c:order val="\d"\/>/g, (_m, d) => `<c:idx val="${Number(d) + 1}"/><c:order val="${Number(d) + 1}"/>`)}<c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/></c:upDownBars><c:axId val="3"/><c:axId val="4"/></c:stockChart>
${axes(1, 2, 'l')}${axes(3, 4, 'r', '<c:crosses val="max"/>')}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;

	it('draws the volume columns on the primary axis and the candles on the secondary axis', () => {
		const svg = render(volume);
		// Four volume columns and four candles, each candle with a body.
		expect(svg.match(/<g data-chart-series="0" data-chart-point="\d"/g)).toHaveLength(4);
		expect(svg.match(/<g data-chart-series="3" data-chart-point="\d"/g)).toHaveLength(4);
		for (const name of ['Volume', 'Open', 'High', 'Low', 'Close'])
			expect(svg).toContain(`>${name}<`);
		// The price scale (about 9 to 15) labels the right edge; the volume scale the left.
		expect(svg).toContain('>1,500<');
		expect(svg).toContain('>14<');
	});
});

describe('bubble headroom', () => {
	const bubble = (xs: number[], ys: number[], sizes: number[]) => {
		const source = (tag: string, values: number[]) =>
			`<c:${tag}><c:numRef><c:f>S!$A$1</c:f><c:numCache>${num(values)}</c:numCache></c:numRef></c:${tag}>`;
		return `<c:chartSpace ${C}><c:chart><c:plotArea><c:layout/><c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/>${label('P')}${source('xVal', xs)}${source('yVal', ys)}${source('bubbleSize', sizes)}</c:ser><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>
<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
	};

	it('keeps bubbles on the extreme axis values inside the plot', () => {
		const svg = render(bubble([1, 2, 3, 4], [1, 2, 3, 8], [1, 2, 3, 10]));
		const circles = [...svg.matchAll(/<circle cx="([\d.]+)" cy="([\d.]+)" r="([\d.]+)"/g)].map(
			(m) => ({
				cx: Number(m[1]),
				cy: Number(m[2]),
				r: Number(m[3]),
			}),
		);
		expect(circles).toHaveLength(4);
		const lines = [
			...svg.matchAll(/<line x1="([\d.]+)" y1="([\d.]+)" x2="([\d.]+)" y2="([\d.]+)"/g),
		].map((m) => m.slice(1).map(Number) as [number, number, number, number]);
		const left = Math.min(...lines.map((l) => l[0]));
		const right = Math.max(...lines.map((l) => l[2]));
		const top = Math.min(...lines.map((l) => l[1]));
		const bottom = Math.max(...lines.map((l) => l[3]));
		for (const c of circles) {
			expect(c.cx - c.r).toBeGreaterThanOrEqual(left - 0.01);
			expect(c.cx + c.r).toBeLessThanOrEqual(right + 0.01);
			expect(c.cy - c.r).toBeGreaterThanOrEqual(top - 0.01);
			expect(c.cy + c.r).toBeLessThanOrEqual(bottom + 0.01);
		}
	});
});
