import { expect, it } from 'vitest';
import { parseXml, NS, first, children } from '../../xml/index';
import { writeDiagramFormatDefinition } from './format';
import { buildFabricatedColorsXml, buildFabricatedQuickStyleXml } from '../write/fabricated-styles';
import { writeEditedDiagramData } from '../write/data-edit';
import { writeInsertedDiagramData } from '../write/inserted';
import { parseDiagramDataModel } from '../data-model';

it('changes colors while preserving unknown labels, XML, attributes and theme references', () => {
	const xml = buildFabricatedColorsXml('colorful1')
		.replace('<dgm:colorsDef ', '<dgm:colorsDef xmlns:x="urn:test" x:kept="yes" ')
		.replace('</dgm:colorsDef>', '<dgm:extLst><x:kept val="42"/></dgm:extLst></dgm:colorsDef>');
	const result = writeDiagramFormatDefinition(xml, { kind: 'colors', value: 'monochromatic2' });
	expect(result).toContain('x:kept="yes"');
	expect(result).toContain('<x:kept val="42"');
	expect(result).toContain('urn:ooxml-core/colors/monochromatic2');
	const root = parseXml(result).documentElement;
	const node = children(root, 'styleLbl', NS.dgm).find(
		(label) => label.getAttribute('name') === 'node1',
	)!;
	expect(first(first(node, 'fillClrLst', NS.dgm), 'schemeClr', NS.a)?.getAttribute('val')).toBe(
		'accent2',
	);
	expect(result).not.toContain('srgbClr');
});

it('changes supported style matrix indices without removing authored 3D, fonts or extensions', () => {
	const xml = buildFabricatedQuickStyleXml('flat')
		.replace('<dgm:styleDef ', '<dgm:styleDef xmlns:x="urn:test" ')
		.replace('</dgm:styleDef>', '<dgm:extLst><x:future/></dgm:extLst></dgm:styleDef>');
	const result = writeDiagramFormatDefinition(xml, { kind: 'style', value: 'intense' });
	expect(result).toContain('<x:future');
	expect(result).toContain('orthographicFront');
	const label = children(parseXml(result).documentElement, 'styleLbl', NS.dgm).find(
		(node) => node.getAttribute('name') === 'node1',
	)!;
	const style = first(label, 'style', NS.dgm);
	expect(first(style, 'fillRef', NS.a)?.getAttribute('idx')).toBe('3');
	expect(first(style, 'effectRef', NS.a)?.getAttribute('idx')).toBe('2');
	expect(first(style, 'fontRef', NS.a)?.getAttribute('idx')).toBe('minor');
});

it('updates document definition identities without discarding node overrides or unrelated data XML', () => {
	const input = writeInsertedDiagramData({
		nodes: [{ id: 'a', text: 'Alpha' }],
		layoutId: 'lo',
		colorsId: 'old-colors',
		quickStyleId: 'old-style',
	});
	const node = parseDiagramDataModel(input).nodes[0]!;
	const source = input.replace(
		'</dgm:dataModel>',
		'<dgm:whole><dgm:spPr/></dgm:whole></dgm:dataModel>',
	);
	const result = writeEditedDiagramData(source, [{ id: node.modelId, text: 'Alpha' }], {
		colorsId: 'new-colors',
		quickStyleId: 'new-style',
	});
	const model = parseDiagramDataModel(result);
	expect(model.points.find((point) => point.type === 'doc')).toMatchObject({
		colorsTypeId: 'new-colors',
		quickStyleTypeId: 'new-style',
	});
	expect(result).toContain('<dgm:whole>');
	expect(model.nodes[0]?.text).toBe('Alpha');
});
