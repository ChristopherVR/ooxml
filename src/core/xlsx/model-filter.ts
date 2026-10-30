import type { DifferentialStyle } from './model';

/** `ST_FilterOperator`: `equal` is the default of a `customFilter`. */
export type FilterOperator =
	| 'equal'
	| 'lessThan'
	| 'lessThanOrEqual'
	| 'notEqual'
	| 'greaterThanOrEqual'
	| 'greaterThan';

export type DateGroupLevel = 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second';

/** `dateGroupItem`: every date inside one year, month, day, hour, minute or second. */
export interface DateGroupItem {
	year: number;
	month?: number;
	day?: number;
	hour?: number;
	minute?: number;
	second?: number;
	grouping: DateGroupLevel;
}

export interface CustomFilterCondition {
	operator: FilterOperator;
	/** Text or number as typed; `*`, `?` and `~` wildcards apply to `equal` / `notEqual` text. */
	value: string;
}

/** `customFilters`: one or two conditions joined by AND (`and`) or OR. */
export interface CustomFilters {
	and?: true;
	conditions: CustomFilterCondition[];
}

/** `top10`: the top or bottom `value` items (or percent) of the column. */
export interface Top10Filter {
	top: boolean;
	percent: boolean;
	value: number;
	/** The cut-off Excel cached when it last applied the filter. */
	filterValue?: number;
}

export type DynamicFilterType =
	| 'null'
	| 'aboveAverage'
	| 'belowAverage'
	| 'tomorrow'
	| 'today'
	| 'yesterday'
	| 'nextWeek'
	| 'thisWeek'
	| 'lastWeek'
	| 'nextMonth'
	| 'thisMonth'
	| 'lastMonth'
	| 'nextQuarter'
	| 'thisQuarter'
	| 'lastQuarter'
	| 'nextYear'
	| 'thisYear'
	| 'lastYear'
	| 'yearToDate'
	| 'Q1'
	| 'Q2'
	| 'Q3'
	| 'Q4'
	| 'M1'
	| 'M2'
	| 'M3'
	| 'M4'
	| 'M5'
	| 'M6'
	| 'M7'
	| 'M8'
	| 'M9'
	| 'M10'
	| 'M11'
	| 'M12';

export interface DynamicFilter {
	type: DynamicFilterType;
	/** Cached average or start serial. */
	value?: number;
	maxValue?: number;
}

/** `colorFilter`: cells whose fill (or font, `cellColor` false) colour matches the differential format. */
export interface ColorFilter {
	/** Filter on the cell fill (the default) rather than the font colour. */
	cellColor: boolean;
	dxfId?: number;
	/** The differential format `dxfId` named, resolved when the sheet was read or authored. */
	style?: DifferentialStyle;
}

/** `iconFilter`: cells showing one icon of a conditional-format icon set. */
export interface IconFilter {
	iconSet: string;
	/** Zero-based icon within the set; absent means "No Cell Icon". */
	iconId?: number;
}
