import { describe, expect, it } from 'vitest';
import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { projectChartCompatibilityTree } from './chart-compat-tree';
import { readOriginalChartPart } from './chart-part-repair';

const attributes =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const source = (content: string) => `<c:chartSpace ${attributes}>${content}</c:chartSpace>`;

describe('a chart compatibility tree projected from the shared DOM', () => {
	const parser = new PptxRuntimeDependencyFactory().createParser();

	it('matches the package parser for grouped nodes, whitespace, attributes and empty elements', () => {
		const xml = source(
			`<c:chart>
				<c:spPr/>
				<c:title label="a&amp;b">
					<a:t xml:space="preserve">  Q&amp;A  </a:t>
					<c:separator> / </c:separator>
				</c:title>
				<c:legend pos="b"/>
				<c:legend pos="t"/>
				<c:plotArea><c:ser><c:idx val="0001"/><c:order val="1"/></c:ser></c:plotArea>
			</c:chart>`,
		);
		const projected = projectChartCompatibilityTree(parser, xml, readOriginalChartPart(xml));
		expect(projected).toBeDefined();
		expect(JSON.stringify(projected)).toBe(JSON.stringify(parser.parse(xml)));
	});

	it.each([
		'<c:chart title="One\tTwo"/>',
		'<c:chart><![CDATA[foo]]></c:chart>',
		'<c:chart><c:legend/><!-- comment --></c:chart>',
		'<c:chart><constructor/></c:chart>',
		'<c:chart><c:plotArea val="&#xD;"/></c:chart>',
	])('defers to the object parser for %s', (body) => {
		const xml = source(body);
		expect(projectChartCompatibilityTree(parser, xml, readOriginalChartPart(xml))).toBeUndefined();
	});
});
