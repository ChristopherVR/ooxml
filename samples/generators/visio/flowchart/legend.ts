/** The legend and the reading notes in the column right of the lanes. */
import { polyline } from '../lib/geometry';
import type { PageBuilder } from '../lib/page';
import { cell } from '../lib/xml';

export function drawLegend(page: PageBuilder): void {
	const left = 14.3;
	const legend = { layers: ['Legend'] };
	page.note('Heading', 'Legend', { left, y: 9.33, w: 2.1 });
	const entries: [string, string][] = [
		['Start or end', 'Start or end'],
		['Step', 'Step done by one role'],
		['Decision', 'Decision'],
		['Document', 'Document or notice'],
		['Data', 'Data in a system'],
		['Off-page reference', 'Off-page reference'],
	];
	entries.forEach(([master, caption], index) => {
		const y = 8.85 - index * 0.5;
		const h = master === 'Decision' ? 0.36 : 0.3;
		page.drop(master, {
			x: left + 0.3,
			y,
			w: master === 'Off-page reference' ? 0.42 : 0.6,
			h,
			...legend,
			// The stadium's radius is a formula of the height; its cached value is set for this size.
			...(master === 'Start or end' ? { cells: [cell('Rounding', h / 2, 'Inh')] } : {}),
		});
		page.note('Caption', caption, { left: left + 0.75, y, w: 1.4 });
	});
	const lines: [string, string][] = [
		['Sequence', 'Sequence of work'],
		['Message', 'Document or notice sent'],
	];
	lines.forEach(([style, caption], index) => {
		const y = 5.85 - index * 0.5;
		page.draw(
			{
				style,
				paths: [
					polyline([
						[0, 0.5],
						[1, 0.5],
					]),
				],
			},
			{ x: left + 0.3, y, w: 0.6, h: 0.1, ...legend },
		);
		page.note('Caption', caption, { left: left + 0.75, y, w: 1.4 });
	});
	page.block('Hairline', { left, bottom: 4.95, w: 2.1, h: 0.012 });
	page.note('Heading', 'Reading the chart', { left, y: 4.68, w: 2.1 });
	page.paragraph(
		'Each lane is a role. A step sits in the lane of the role that does the work, and in the phase in which it happens.',
		{ left, top: 4.45, w: 2.1, h: 0.6 },
	);
	page.paragraph(
		'Steps and decisions are linked to the process register. The grey line under a step names the system it is recorded in, read from the shape data.',
		{ left, top: 3.71, w: 2.1, h: 0.75 },
	);
	page.paragraph('Target lead time: two working days from order to collection.', {
		left,
		top: 2.82,
		w: 2.1,
		h: 0.32,
	});
	page.note('Caption', 'Operations handbook, chapter 4', {
		left,
		y: 2.28,
		w: 2.1,
		link: {
			address: 'https://kestrelline.example/handbook/order-to-delivery',
			description: 'Operations handbook, chapter 4',
		},
	});
}
