/** Page 2: the Greywharf tidal energy pilot as intervals and milestones on a two-year axis. */
import type { Drawing } from '../lib/drawing';
import { PAGE } from '../lib/furniture';
import { rectangle } from '../lib/geometry';
import type { Node, PageBuilder } from '../lib/page';
import { HUES, RUN } from '../lib/palette';
import { SIDE } from '../lib/shape';
import { ORGANISATION_PAGE, TIMELINE_PAGE } from './organisation';
import { DEPENDENCIES, MILESTONES, ROWS } from './plan';
import { BAR_TOP, workstream } from './stencil';

/** The axis: 24 months from January 2026, left to right. */
const AXIS = { left: 2.9, right: 16.2, months: 24, firstYear: 2026 };
const MONTH = (AXIS.right - AXIS.left) / AXIS.months;
const monthX = (month: number) => AXIS.left + month * MONTH;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** A date as months from the start of the axis (days counted as thirtieths of a month). */
function monthsOf(date: string): number {
	const [year = 0, month = 1, day = 1] = date.split('-').map(Number);
	return (year - AXIS.firstYear) * 12 + (month - 1) + (day - 1) / 30;
}
/** `2026-03-15` as `15 Mar 2026`, the text Visio shows for the picture `{{d MMM yyyy}}`. */
function shown(date: string): string {
	const [year = 0, month = 1, day = 1] = date.split('-').map(Number);
	return `${day} ${MONTHS[month - 1]} ${year}`;
}

const ROW_TOP = 7.3;
const ROW_STEP = 0.72;
const BAR_HEIGHT = 0.36;

const GRID_BOTTOM = ROW_TOP - (ROWS.length - 1) * ROW_STEP - 0.42;

function drawAxis(page: PageBuilder): void {
	const axis = { layers: ['Axis'] };
	for (let year = 0; year < 2; year++)
		page.draw(
			{ style: 'Axis year', paths: [rectangle()] },
			{
				x: monthX(year * 12 + 6),
				y: 9.2,
				w: MONTH * 12,
				h: 0.3,
				text: String(AXIS.firstYear + year),
				...axis,
			},
		);
	for (let quarter = 0; quarter < 8; quarter++) {
		page.draw(
			{ style: 'Axis quarter', paths: [rectangle()] },
			{
				x: monthX(quarter * 3 + 1.5),
				y: 8.92,
				w: MONTH * 3,
				h: 0.26,
				text: `Q${(quarter % 4) + 1}`,
				...axis,
			},
		);
		// A hairline down the grid at the start of each quarter.
		page.block('Hairline', {
			left: monthX(quarter * 3) - 0.006,
			bottom: GRID_BOTTOM,
			w: 0.012,
			h: 8.79 - GRID_BOTTOM,
		});
	}
	page.block('Hairline', {
		left: AXIS.right - 0.006,
		bottom: GRID_BOTTOM,
		w: 0.012,
		h: 8.79 - GRID_BOTTOM,
	});
	const today = monthsOf('2026-01-15');
	page.block('Today', {
		left: monthX(today) - 0.01,
		bottom: GRID_BOTTOM,
		w: 0.02,
		h: 8.79 - GRID_BOTTOM,
	});
	page.note('Today label', 'Status date: 15 Jan 2026', {
		left: monthX(today) + 0.06,
		y: GRID_BOTTOM + 0.12,
		w: 2,
		h: 0.2,
	});
}

function drawRows(page: PageBuilder): Map<string, Node> {
	const bars = new Map<string, Node>();
	page.note('Row label', 'Milestones', { left: PAGE.left, y: 8.3, w: 2.1 });
	MILESTONES.forEach((milestone) => {
		page.drop('Milestone', {
			x: monthX(monthsOf(milestone.date)),
			y: 8.36,
			text: [
				{ text: `${milestone.name}\n`, style: RUN.name },
				{
					text: shown(milestone.date),
					field: 'Prop.Date',
					picture: '{{d MMM yyyy}}',
					date: milestone.date,
					style: RUN.small,
				},
			],
			data: { Date: milestone.date, Owner: milestone.owner },
			layers: ['Milestones'],
		});
	});
	ROWS.forEach((row, index) => {
		const y = ROW_TOP - index * ROW_STEP;
		page.block('Hairline', {
			left: PAGE.left,
			bottom: y + ROW_STEP / 2,
			w: AXIS.right - PAGE.left,
			h: 0.012,
		});
		page.note('Row label', row.name, { left: PAGE.left, y, w: 2.1 });
		for (const task of row.tasks) {
			const from = monthX(monthsOf(task.start));
			const to = monthX(monthsOf(task.finish));
			bars.set(
				task.key,
				page.drop('Interval', {
					x: (from + to) / 2,
					y,
					w: to - from,
					h: BAR_HEIGHT,
					text: task.name,
					cells: workstream(HUES[row.hue]),
					data: {
						Workstream: row.name,
						Owner: task.owner,
						Start: task.start,
						Finish: task.finish,
						Budget: task.budget,
						Status: task.status ?? 'Planned',
					},
					layers: ['Work'],
				}),
			);
		}
	});
	page.block('Hairline', {
		left: PAGE.left,
		bottom: GRID_BOTTOM,
		w: AXIS.right - PAGE.left,
		h: 0.012,
	});
	return bars;
}

function drawDependencies(page: PageBuilder, bars: ReadonlyMap<string, Node>): void {
	for (const [from, to, how] of DEPENDENCIES) {
		const source = bars.get(from);
		const target = bars.get(to);
		if (!source || !target) throw new Error(`No bar ${from} or ${to}`);
		const ports =
			how === 'over'
				? ([source.port(SIDE.right), target.port(BAR_TOP)] as const)
				: ([source.port(SIDE.bottom), target.port(SIDE.left)] as const);
		page.connect('Dependency', ports[0], ports[1], { layers: ['Dependencies'] });
	}
}

function drawNotes(page: PageBuilder): void {
	const top = GRID_BOTTOM - 0.5;
	page.note('Heading', 'Greywharf tidal energy pilot', { left: PAGE.left, y: top, w: 5 });
	page.paragraph(
		'One 500 kW turbine on a gravity foundation in the Greywharf narrows, built to measure yield, loads and the response of seabirds and seals over a first winter.',
		{ left: PAGE.left, top: top - 0.24, w: 5, h: 0.5 },
	);
	page.note('Heading', 'Reading the timeline', { left: 6.3, y: top, w: 5 });
	page.paragraph(
		'A bar is a piece of work in its workstream, a diamond is a dated milestone, and an arrow means the work at its tail must finish first. Start, finish, owner and budget are shape data on each bar; the date under a milestone is a field.',
		{ left: 6.3, top: top - 0.24, w: 5, h: 0.62 },
	);
	page.note('Heading', 'Budget', { left: 12.0, y: top, w: 4.4 });
	const total = ROWS.flatMap((row) => row.tasks).reduce((sum, task) => sum + task.budget, 0);
	page.note('Caption', `${(total / 1000).toFixed(2)} million across ${ROWS.length} workstreams.`, {
		left: 12.0,
		y: top - 0.31,
		w: 4.4,
	});
	page.note('Fact', 'Who is who: organisation chart', {
		left: 12.0,
		y: top - 0.61,
		w: 4.4,
		link: { page: ORGANISATION_PAGE, description: 'The organisation chart' },
	});
}

/** Draws the page; the milestone date fields are held back like the other fields. */
export async function drawTimelinePage(drawing: Drawing): Promise<PageBuilder> {
	const page = drawing.page(TIMELINE_PAGE);
	page.deferFields = true;
	drawAxis(page);
	const bars = drawRows(page);
	drawDependencies(page, bars);
	drawNotes(page);
	await drawing.place(page);
	return page;
}
