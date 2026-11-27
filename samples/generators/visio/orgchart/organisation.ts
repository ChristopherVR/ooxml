/** Page 1: the organisation chart, a tree from the director to the teams. */
import type { Drawing } from '../lib/drawing';
import { PAGE } from '../lib/furniture';
import { ellipse } from '../lib/geometry';
import type { Node, PageBuilder } from '../lib/page';
import { type HueName, FOG, HUES, RUN } from '../lib/palette';
import type { RecordsetSpec } from '../lib/recordset';
import { SIDE } from '../lib/shape';
import { cell, colorCell } from '../lib/xml';
import { type Post, DEPARTMENTS, DIRECTOR, DIRECTORATE, email } from './staff';
import { POSITION, TEAM_POINT, department } from './stencil';

export const ORGANISATION_PAGE = 'Organisation';
export const TIMELINE_PAGE = 'Greywharf pilot timeline';

/** Eight team columns across the page; a head sits over its two teams, the director over all. */
const columnX = (team: number) => PAGE.left + 1.075 + team * 1.95;
const ROWS = { director: 9.1, head: 7.9, lead: 6.7, member: 5.72, step: 0.82 };
const MEMBER = { w: 1.55, h: 0.66, indent: 0.1 };

interface Placed {
	post: Post;
	department: string;
	node: Node;
}

function drawTree(page: PageBuilder): Placed[] {
	const placed: Placed[] = [];
	const lines = { layers: ['Reporting lines'], square: true };
	const card = (
		post: Post,
		unit: { name: string; hue: HueName },
		at: { x: number; y: number; w?: number; h?: number },
		layers: readonly string[],
		link?: boolean,
	): Node => {
		const node = page.drop('Position', {
			...at,
			text: [
				{ text: `${post.name}\n`, style: RUN.name },
				{ text: post.title, field: 'Prop.Title', style: RUN.small },
			],
			childCells: department(HUES[unit.hue]),
			// A vacant post is drawn hollow: a dashed outline on a grey card.
			...(post.status === 'Vacant'
				? { cells: [cell('LinePattern', 2), colorCell('FillForegnd', FOG)] }
				: {}),
			layers,
			...(link
				? { link: { page: TIMELINE_PAGE, description: 'The Greywharf pilot timeline' } }
				: {}),
		});
		placed.push({ post, department: unit.name, node });
		return node;
	};
	const director = card(
		DIRECTOR,
		DIRECTORATE,
		{ x: (columnX(0) + columnX(7)) / 2, y: ROWS.director },
		['Leadership'],
	);
	DEPARTMENTS.forEach((unit, index) => {
		const [left, right] = [columnX(index * 2), columnX(index * 2 + 1)];
		const head = card(
			unit.head,
			unit,
			{ x: (left + right) / 2, y: ROWS.head },
			[unit.name, 'Leadership'],
			unit.name === 'Marine Energy',
		);
		page.connect('Reporting line', director.port(SIDE.bottom), head.port(SIDE.top), lines);
		unit.teams.forEach((team, at) => {
			const x = at === 0 ? left : right;
			const lead = card(team.lead, unit, { x, y: ROWS.lead }, [unit.name]);
			page.connect('Reporting line', head.port(SIDE.bottom), lead.port(SIDE.top), lines);
			team.members.forEach((post, row) => {
				const member = card(
					post,
					unit,
					{
						x: x + (POSITION.w - MEMBER.w) / 2,
						y: ROWS.member - row * ROWS.step,
						w: MEMBER.w,
						h: MEMBER.h,
					},
					[unit.name],
				);
				page.connect('Reporting line', lead.port(TEAM_POINT), member.port(SIDE.left), lines);
			});
		});
	});
	return placed;
}

/** The panels under the tree: the departments, how to read the chart, and what is changing. */
function drawPanels(page: PageBuilder, placed: readonly Placed[]): void {
	const notes = { layers: ['Notes'] };
	const top = 3.72;
	page.block('Hairline', { left: PAGE.left, bottom: 4.02, w: PAGE.right - PAGE.left, h: 0.012 });
	page.note('Heading', 'Departments', { left: PAGE.left, y: top, w: 4 });
	[DIRECTORATE, ...DEPARTMENTS].forEach((unit, index) => {
		const y = top - 0.33 - index * 0.3;
		const posts = placed.filter((entry) => entry.department === unit.name).length;
		page.draw(
			{ style: `Swatch ${unit.hue}`, paths: [ellipse(0.5, 0.5, 0.5, 0.5)] },
			{ x: PAGE.left + 0.09, y, w: 0.18, h: 0.18, ...notes },
		);
		page.note('Fact', unit.name, { left: PAGE.left + 0.32, y, w: 2.2 });
		page.note('Caption', `${posts} ${posts === 1 ? 'post' : 'posts'}`, {
			left: PAGE.left + 2.6,
			y,
			w: 1.2,
		});
	});
	const middle = 5.9;
	page.note('Heading', 'Reading the chart', { left: middle, y: top, w: 4.6 });
	page.paragraph(
		'Every card is a post. The coloured edge is its department, and a dashed card is a post that is not filled yet.',
		{ left: middle, top: top - 0.24, w: 4.6, h: 0.32 },
	);
	page.paragraph(
		'Cards are linked to the staff directory, a data recordset kept in the drawing: name, title, department, location, start date and working time. The title under a name is a field that reads it from there.',
		{ left: middle, top: top - 0.7, w: 4.6, h: 0.48 },
	);
	const right = 11.4;
	page.note('Heading', 'Changes this quarter', { left: right, y: top, w: 5 });
	const changes = [
		'Coastal ecologist, Sediment Dynamics: recruiting, start planned for 2 February 2026.',
		'Yuki Tanabe is seconded to the Greywharf pilot until first power.',
		'Zofia Marchetti joined Research Computing at three days a week.',
	];
	changes.forEach((text, index) => {
		page.note('Caption', text, { left: right, y: top - 0.31 - index * 0.28, w: 5 });
	});
	page.note('Fact', 'Greywharf pilot timeline, next sheet', {
		left: right,
		y: top - 1.27,
		w: 5,
		link: { page: TIMELINE_PAGE, description: 'The Greywharf pilot timeline' },
	});
	page.note('Caption', 'People directory', {
		left: right,
		y: top - 1.55,
		w: 5,
		link: {
			address: 'https://saltmarsh.example/people',
			description: 'Saltmarsh Institute people directory',
		},
	});
}

/**
 * Draws the page and links every card to its row of the staff directory. The title fields are
 * still held back; the caller applies them last.
 */
export async function drawOrganisationPage(drawing: Drawing): Promise<PageBuilder> {
	const page = drawing.page(ORGANISATION_PAGE);
	page.deferFields = true;
	const placed = drawTree(page);
	drawPanels(page, placed);
	await drawing.place(page);
	const directory: RecordsetSpec = {
		name: 'Staff directory',
		columns: [
			{ name: 'StaffId', label: 'Staff number', type: 'string' },
			{ name: 'Name', label: 'Name', type: 'string' },
			{ name: 'Title', label: 'Title', type: 'string' },
			{ name: 'Department', label: 'Department', type: 'string' },
			{ name: 'Location', label: 'Location', type: 'string' },
			{ name: 'Email', label: 'Email', type: 'string' },
			{ name: 'Started', label: 'Started', type: 'date' },
			{ name: 'FTE', label: 'Full-time equivalent', type: 'number' },
			{ name: 'Status', label: 'Status', type: 'string' },
		],
		rows: placed.map(({ post, department: unit }, index) => [
			`SI-${String(101 + index)}`,
			post.name,
			post.title,
			unit,
			post.location,
			email(post),
			post.started,
			String(post.fte ?? 1),
			post.status ?? 'In post',
		]),
		primaryKey: 'StaffId',
	};
	await drawing.linkData(
		page,
		directory,
		placed.map((entry, index) => [entry.node.id, index + 1] as const),
	);
	return page;
}
