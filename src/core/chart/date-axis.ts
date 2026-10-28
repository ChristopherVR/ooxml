const DAY_MS = 86_400_000;

export function excelSerialToDate(serial: number, date1904 = false): Date {
	const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, serial < 60 ? 31 : 30);
	return new Date(epoch + serial * DAY_MS);
}

function addUnit(date: Date, amount: number, unit: 'days' | 'months' | 'years'): Date {
	const next = new Date(date);
	if (unit === 'days') {
		next.setUTCDate(next.getUTCDate() + amount);
	} else if (unit === 'months') {
		next.setUTCMonth(next.getUTCMonth() + amount);
	} else {
		next.setUTCFullYear(next.getUTCFullYear() + amount);
	}
	return next;
}

export function calendarTicks(
	min: number,
	max: number,
	date1904: boolean,
	unit: 'days' | 'months' | 'years',
	step: number,
): number[] {
	const result: number[] = [];
	let date = excelSerialToDate(min, date1904);
	for (
		let guard = 0;
		guard < 1000 && date.getTime() <= excelSerialToDate(max, date1904).getTime();
		guard++
	) {
		result.push(min + (date.getTime() - excelSerialToDate(min, date1904).getTime()) / DAY_MS);
		date = addUnit(date, step, unit);
	}
	return result;
}
