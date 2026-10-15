// The calendar a date picker content control opens: month navigation, a day grid and "Today".
import { localeOf, translate } from './localization';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';

const sameDay = (a: Date, b: Date): boolean =>
	a.getFullYear() === b.getFullYear() &&
	a.getMonth() === b.getMonth() &&
	a.getDate() === b.getDate();

/** The 6 x 7 grid of days that shows `month`, weeks starting on Sunday. */
export function calendarDays(year: number, month: number): Date[] {
	const first = new Date(year, month, 1);
	const start = new Date(year, month, 1 - first.getDay());
	return Array.from(
		{ length: 42 },
		(_, i) => new Date(start.getFullYear(), start.getMonth(), start.getDate() + i),
	);
}

/**
 * Opens the calendar under `anchor`. `selected` marks the current value and sets the month shown;
 * `language` is a BCP 47 tag for month and weekday names.
 */
export function openCalendarPopover(
	anchor: HTMLElement,
	selected: Date | undefined,
	language: string,
	choose: (date: Date) => void,
): void {
	const locale = localeOf(anchor);
	const today = new Date();
	let shown = new Date((selected ?? today).getFullYear(), (selected ?? today).getMonth(), 1);
	const names = (options: Intl.DateTimeFormatOptions, date: Date) => {
		try {
			return new Intl.DateTimeFormat(language, options).format(date);
		} catch {
			return new Intl.DateTimeFormat('en-US', options).format(date);
		}
	};
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover sdt-calendar';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', translate(locale, 'Date picker'));
	const head = document.createElement('div');
	head.className = 'sdt-calendar-head';
	const previous = document.createElement('button');
	previous.type = 'button';
	previous.textContent = '‹';
	previous.setAttribute('aria-label', translate(locale, 'Previous month'));
	const next = document.createElement('button');
	next.type = 'button';
	next.textContent = '›';
	next.setAttribute('aria-label', translate(locale, 'Next month'));
	const title = document.createElement('span');
	title.setAttribute('aria-live', 'polite');
	head.append(previous, title, next);
	const grid = document.createElement('div');
	grid.className = 'sdt-calendar-grid';
	grid.setAttribute('role', 'grid');
	const foot = document.createElement('div');
	foot.className = 'sdt-calendar-foot';
	const todayButton = document.createElement('button');
	todayButton.type = 'button';
	todayButton.textContent = translate(locale, 'Today');
	foot.append(todayButton);
	pop.append(head, grid, foot);
	for (const control of [previous, next, todayButton])
		control.addEventListener('mousedown', (event) => event.preventDefault());
	const done = (date: Date) => {
		choose(date);
		closeRibbonPopover();
	};
	const paint = () => {
		title.textContent = names({ month: 'long', year: 'numeric' }, shown);
		const week = Array.from({ length: 7 }, (_, i) => {
			const cell = document.createElement('span');
			cell.textContent = names({ weekday: 'narrow' }, new Date(2024, 0, 7 + i));
			cell.setAttribute('role', 'columnheader');
			return cell;
		});
		const days = calendarDays(shown.getFullYear(), shown.getMonth()).map((date) => {
			const day = document.createElement('button');
			day.type = 'button';
			day.textContent = String(date.getDate());
			day.dataset.date = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
			day.setAttribute('role', 'gridcell');
			day.setAttribute('aria-label', names({ dateStyle: 'full' }, date));
			day.classList.toggle('outside', date.getMonth() !== shown.getMonth());
			day.classList.toggle('today', sameDay(date, today));
			const isSelected = Boolean(selected && sameDay(date, selected));
			day.classList.toggle('selected', isSelected);
			if (isSelected) day.setAttribute('aria-selected', 'true');
			day.addEventListener('mousedown', (event) => event.preventDefault());
			day.addEventListener('click', () => done(date));
			return day;
		});
		grid.replaceChildren(...week, ...days);
	};
	previous.addEventListener('click', () => {
		shown = new Date(shown.getFullYear(), shown.getMonth() - 1, 1);
		paint();
	});
	next.addEventListener('click', () => {
		shown = new Date(shown.getFullYear(), shown.getMonth() + 1, 1);
		paint();
	});
	todayButton.addEventListener('click', () => done(new Date()));
	paint();
	mountPopover(anchor, pop, anchor);
}
