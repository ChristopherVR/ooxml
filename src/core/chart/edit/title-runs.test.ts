import { expect, it } from 'vitest';
import { newChartSpace } from './new-chart';
import { setChartTitleRuns } from './title-runs';
import { chartTextBodyFromXml } from '../write-shape';
import { writeChartSpace } from '../write-space';
const rich =
	'<a:bodyPr/><a:p><a:pPr marL="15"/><a:r><a:rPr b="1"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Producer"/></a:rPr><a:t>First</a:t></a:r><a:r><a:rPr i="1"/><a:t>Second</a:t></a:r><a:endParaRPr lang="fr-FR"/></a:p>';
it('patches explicit runs while keeping theme colours and paragraph/font metadata', () => {
	const space = newChartSpace(
		{ categories: ['A'], series: [{ name: 'S', values: [1] }] },
		{ kind: 'bar' },
	);
	space.title = { tx: { rich: chartTextBodyFromXml(rich) }, text: 'FirstSecond' };
	const changed = setChartTitleRuns(space, [
		{ text: 'First edited' },
		{ text: 'Second', format: { sizePt: 23 }, inherit: ['italic'] },
	]);
	const xml = writeChartSpace(changed);
	expect(xml).toContain('typeface="Producer"');
	expect(xml).toContain('schemeClr val="accent1"');
	expect(xml).toContain('marL="15"');
	expect(xml).toContain('lang="fr-FR"');
	expect(xml).toContain('sz="2300"');
	expect(xml).not.toContain('i="1"');
	expect(space.title.tx!.rich!.text).toBe('FirstSecond');
});
it('adds and removes runs while keeping the last paragraph properties', () => {
	const space = newChartSpace(
		{ categories: ['A'], series: [{ name: 'S', values: [1] }] },
		{ kind: 'bar' },
		{ title: 'One' },
	);
	const added = setChartTitleRuns(space, [{ text: 'A' }, { text: 'B', format: { bold: true } }]);
	expect(added.title!.text).toBe('AB');
	const removed = setChartTitleRuns(added, [{ text: 'Only' }]);
	expect(removed.title!.text).toBe('Only');
	expect(removed.title!.tx!.rich!.paragraphs.flatMap((p) => p.runs)).toHaveLength(1);
	expect(() => setChartTitleRuns(space, [{ text: 'X', format: { sizePt: Infinity } }])).toThrow();
});
