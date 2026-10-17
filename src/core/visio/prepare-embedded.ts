import { NS } from '../xml/index';
import { readEmbeddedChart } from '../xlsx/ui/dialogs/grid-chart';
import { VISIO_CHART_OBJECT_MAX_BYTES } from './edit-chart-object-commands';
import type { VisioEmbeddedObject } from './model';
import type { VisioPackage } from './package';
import { children, type Report } from './sheet';

/**
 * An embedded Excel workbook that holds a chart (Insert > Chart in Visio, or this editor): the
 * Foreign shape's ForeignData is an Object whose relationship is a package. Anything else embedded
 * (a Word document, a PDF) is not read and the shape is reported as unsupported, as before.
 */
export async function readEmbeddedChartObject(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
): Promise<VisioEmbeddedObject | undefined> {
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel || rel.mode !== 'Internal' || rel.type !== `${NS.r}/package`) return undefined;
	if (pkg.getPartByteLength(rel.target) > VISIO_CHART_OBJECT_MAX_BYTES) {
		report('embedded-object-limit', 'An embedded object is larger than 8 MiB and was not read.');
		return undefined;
	}
	const bytes = await pkg.readBytes(rel.target);
	const chart = await readEmbeddedChart(bytes);
	if (!chart) return undefined;
	const objectType = Number(foreignData.getAttribute('ObjectType'));
	return {
		...(Number.isFinite(objectType) ? { objectType } : {}),
		bytes,
		chart: chart.view,
	};
}
