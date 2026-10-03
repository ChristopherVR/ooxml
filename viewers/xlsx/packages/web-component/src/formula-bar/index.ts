// The formula bar: Name Box, cancel / enter / fx buttons and the expandable formula input.
import type { EditorContext } from '../context.js';
import { ensureStyle, h } from '../grid/dom.js';
import { createFormulaInput } from './formula-input.js';
import { createNameBox } from './name-box.js';
import { FORMULA_BAR_CSS } from './styles.js';

export {
	absoluteReference,
	nameForRange,
	parseSheetReference,
	resolveNameBox,
} from './name-resolve.js';

export function mountFormulaBar(ctx: EditorContext, container: HTMLElement): () => void {
	const doc = container.ownerDocument;
	ensureStyle(ctx.root, 'xfb', FORMULA_BAR_CSS);
	const bar = h(doc, 'div', 'xfb', {
		part: 'formula-bar',
		role: 'group',
		'aria-label': ctx.t('Formula Bar'),
	});
	const nameBox = createNameBox(ctx, doc);
	bar.append(nameBox.root);
	const formula = createFormulaInput(ctx, doc, bar);
	container.append(bar);
	const refresh = () => {
		nameBox.refresh();
		formula.refresh();
	};
	const offSelection = ctx.selection.onChange(refresh);
	const offModel = ctx.onModelChange(refresh);
	return () => {
		offSelection();
		offModel();
		formula.destroy();
		nameBox.destroy();
		bar.remove();
	};
}
