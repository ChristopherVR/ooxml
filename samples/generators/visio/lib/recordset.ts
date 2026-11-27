/** Data recordsets: the table a drawing keeps, imported through the library. */
import type { VisioDataColumn } from '../../../../src/core/visio/data-recordsets';

/** A table kept in the drawing as a data recordset; its rows are linked to shapes. */
export interface RecordsetSpec {
	name: string;
	columns: readonly VisioDataColumn[];
	/** Text per column: numbers, TRUE/FALSE, ISO dates (yyyy-mm-dd) or strings. */
	rows: readonly (readonly string[])[];
	/** The column that identifies a row. */
	primaryKey: string;
}
