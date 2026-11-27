/** Data recordsets: the table a drawing keeps and the rows part Visio reads it from. */
import type { VisioDataColumn } from '../../../../src/core/visio/data-recordsets';
import { xmlEscape } from '../../lib/package';

/** A table kept in the drawing as a data recordset; its rows are linked to shapes. */
export interface RecordsetSpec {
	name: string;
	columns: readonly VisioDataColumn[];
	/** Text per column: numbers, TRUE/FALSE, ISO dates (yyyy-mm-dd) or strings. */
	rows: readonly (readonly string[])[];
	/** The column that identifies a row. */
	primaryKey: string;
}

const ADO_TYPES = { string: 'string', number: 'float', boolean: 'boolean', date: 'dateTime' };

/**
 * The rows of a recordset as the ADO persisted rowset Visio reads. The library's own part has no
 * `_Visio_RowID_` column, and Visio then never returns from DataRecordset.GetDataRowIDs, so the
 * part is written here: the row number first, then the columns.
 */
export function rowsetXml(spec: RecordsetSpec): string {
	const columns = spec.columns
		.map((column, index) => {
			const length = column.type === 'string' ? ' dt:maxLength="255"' : '';
			return (
				`<s:AttributeType name="c${index + 1}" rs:name="${xmlEscape(column.name)}" rs:number="${index + 2}" rs:nullable="true" rs:write="true">` +
				`<s:datatype dt:type="${ADO_TYPES[column.type]}"${length}/></s:AttributeType>`
			);
		})
		.join('');
	const rows = spec.rows
		.map((values, index) => {
			const cells = values.map((value, at) => {
				const text = spec.columns[at]?.type === 'date' ? `${value}T00:00:00` : value;
				return ` c${at + 1}="${xmlEscape(text)}"`;
			});
			return `<z:row c0="${index + 1}"${cells.join('')}/>`;
		})
		.join('');
	return (
		'<xml xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:rs="urn:schemas-microsoft-com:rowset" xmlns:z="#RowsetSchema">' +
		'<s:Schema id="RowsetSchema"><s:ElementType name="row" content="eltOnly" rs:updatable="true">' +
		'<s:AttributeType name="c0" rs:name="_Visio_RowID_" rs:number="1" rs:nullable="true" rs:maydefer="true" rs:write="true">' +
		'<s:datatype dt:type="int" dt:maxLength="4" rs:precision="0" rs:fixedlength="true"/></s:AttributeType>' +
		`${columns}<s:extends type="rs:rowbase"/></s:ElementType></s:Schema><rs:data>${rows}</rs:data></xml>`
	);
}
