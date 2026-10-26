import type { ChartSpace } from '../model';
import type { ChartManualLayout } from '../manual-layout';
import { chartAxisAt, type ChartAxisTarget } from './axes';
import { assignOrDelete, editChartSpace } from './draft';
import { chartTitleWithText } from './title';
import { editChartLayout } from './layout';
import { editChartTextBody, type ChartTextEdit } from './text';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';

export interface ChartDisplayUnitsLabelEdit {
	text?: string | null;
	layout?: ChartManualLayout | null;
	shape?: ChartShapeEdit | null;
	font?: ChartTextEdit;
}

export function setChartDisplayUnitsLabel(
	space: ChartSpace,
	target: ChartAxisTarget,
	edit: ChartDisplayUnitsLabelEdit | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const axis = chartAxisAt(draft, target);
		if (axis.kind !== 'val') throw new RangeError('Only a value axis has a display-units label');
		if (edit === null) {
			if (axis.displayUnits) delete axis.displayUnits.label;
			return;
		}
		axis.displayUnits ??= {};
		const label = (axis.displayUnits.label ??= {});
		if (edit.text === null) delete label.tx;
		else if (edit.text !== undefined)
			assignOrDelete(
				label,
				'tx',
				chartTitleWithText(label.tx ? { tx: label.tx } : undefined, edit.text, {
					preserveRuns: true,
				}).tx,
			);
		if (edit.layout !== undefined) label.layout = editChartLayout(label.layout, edit.layout);
		if (edit.shape === null) delete label.spPr;
		else if (edit.shape) label.spPr = editChartShapeProperties(label.spPr, edit.shape);
		if (edit.font) {
			label.txPr = editChartTextBody(label.txPr, edit.font);
			if (label.tx?.rich) label.tx.rich = editChartTextBody(label.tx.rich, edit.font);
		}
	});
}
