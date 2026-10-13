import { FootnoteFlow } from './footnote-flow';
import type { ReflowFootnote } from './footnote-columns';
import type { LayoutFootnoteBox, LayoutPageBox } from './result';
/** Reuse single-column reservation and splitting through independent body-column page views. */
export class LegacyFootnoteFlow {
	private readonly flow: FootnoteFlow;
	private readonly views = new WeakMap<LayoutPageBox, Map<number, LayoutPageBox>>();
	constructor(
		report: (message: string) => void,
		separator:
			| ((page: LayoutPageBox, continued: boolean) => LayoutPageBox['footnoteSeparator'])
			| undefined,
		reflow: ReflowFootnote = (note) => note,
	) {
		this.flow = new FootnoteFlow(report, separator, undefined, reflow);
	}
	private view(page: LayoutPageBox, index: number): LayoutPageBox {
		let views = this.views.get(page);
		if (!views) {
			views = new Map();
			this.views.set(page, views);
		}
		let view = views.get(index);
		if (!view) {
			const {
				footnotes: _notes,
				footnoteSeparator: _separator,
				footnoteHeightPx: _height,
				footnoteColumnAreas: _areas,
				...geometry
			} = page;
			view = {
				...geometry,
				widthPx: page.columns[index]!.widthPx,
				marginLeftPx: 0,
				marginRightPx: 0,
				columns: [],
			};
			views.set(index, view);
		}
		return view;
	}
	private sync(page: LayoutPageBox): void {
		const views = this.views.get(page);
		if (!views) return;
		const areas = [...views]
			.filter(([, view]) => view.footnotes?.length)
			.map(([columnIndex, view]) => {
				const column = page.columns[columnIndex]!;
				return {
					columnIndex,
					xPx: column.xPx,
					widthPx: column.widthPx,
					heightPx: view.footnotes!.reduce((sum, note) => sum + note.heightPx, 0),
					continued: !!view.footnotes![0]?.continued,
					...(view.footnoteSeparator ? { separator: view.footnoteSeparator } : {}),
					yPx: 0,
				};
			});
		const height = Math.max(0, ...areas.map((area) => area.heightPx));
		for (const area of areas) area.yPx = height - area.heightPx;
		page.footnoteColumnAreas = areas;
		page.footnoteHeightPx = height;
		page.footnotes = areas.flatMap((area) =>
			this.view(page, area.columnIndex).footnotes!.map((note) => ({
				...note,
				xPx: area.xPx,
				widthPx: area.widthPx,
				columnIndex: area.columnIndex,
				yPx: area.yPx + note.yPx,
			})),
		);
	}
	get hasContinuation(): boolean {
		return this.flow.hasContinuation;
	}
	hold(
		blockId: string,
		notes: LayoutFootnoteBox[],
		bodyMinimum: number,
		freshBodyMinimum = bodyMinimum,
	): void {
		this.flow.hold(blockId, notes, bodyMinimum, freshBodyMinimum);
	}
	reserved(page: LayoutPageBox, columnIndex = 0): number {
		return this.flow.reserved(this.view(page, columnIndex));
	}
	pendingHeight(page: LayoutPageBox, capacity: number, bodyUsed: number, columnIndex = 0): number {
		return this.flow.pendingHeight(this.view(page, columnIndex), capacity, bodyUsed);
	}
	blockedByEarlierColumn(_page: LayoutPageBox, _capacity: number, _bodyBottom: number): boolean {
		return false;
	}
	commit(blockId: string, page: LayoutPageBox, available: number, columnIndex = 0): void {
		if (this.flow.commit(blockId, this.view(page, columnIndex), available)) this.sync(page);
	}
	resume(page: LayoutPageBox, capacity: number, columnIndex?: number): void {
		if (!this.flow.hasContinuation) return;
		const indexes =
			columnIndex === undefined ? page.columns.map((_, index) => index) : [columnIndex];
		for (const index of indexes) {
			if (!this.flow.hasContinuation) break;
			const view = this.view(page, index);
			if (view.footnotes?.length) continue;
			this.flow.resume(view, capacity, undefined, view.widthPx);
		}
		this.sync(page);
	}
}
