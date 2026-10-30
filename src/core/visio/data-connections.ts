import { related, visioXml } from './parts';
import { attribute, children } from './sheet';
import type { VisioPackage } from './package';

/** A data connection saved in the drawing (DataConnections part), as Visio refreshes it. */
export interface VisioDataConnection {
	id: string;
	/** OLE DB connection string; never opened by this package. */
	connectionString: string;
	/** Connection file (.odc) name when the connection came from one. */
	fileName?: string;
}
/** Where an imported table came from; the core derives the connection and command from it. */
export interface VisioDataSource {
	kind: 'workbook' | 'csv';
	/** The file name the user picked; the browser never exposes the folder it lives in. */
	fileName: string;
	/** Worksheet name for a workbook. */
	sheet?: string;
	/** A1 range inside the worksheet, when the import was limited to one. */
	range?: string;
}
export const DATA_CONNECTIONS_TYPE = 'application/vnd.ms-visio.connections+xml';

/**
 * The OLE DB connection string and command Visio writes for the same kind of source. A browser
 * only knows the file name, so Data Source is that name, not a full path: Visio opens it relative
 * to the drawing and asks for the file when it is not found there.
 */
export function visioDataConnectionFor(source: VisioDataSource): {
	connectionString: string;
	command: string;
} {
	const name = source.fileName.replace(/[\\/]/g, '_');
	if (source.kind === 'csv')
		return {
			connectionString: `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.;Extended Properties="text;HDR=Yes;FMT=Delimited"`,
			command: `SELECT * FROM [${name}]`,
		};
	const sheet = source.sheet || 'Sheet1';
	return {
		connectionString: `Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=${name};Mode=Read;Extended Properties="HDR=YES;IMEX=1;"`,
		command: `SELECT * FROM [${sheet}$${source.range ?? ''}]`,
	};
}

/** The DataConnections of the drawing by ID; empty when there is no part. */
export async function readDataConnections(
	pkg: VisioPackage,
	documentPart: string,
): Promise<Map<string, VisioDataConnection>> {
	const result = new Map<string, VisioDataConnection>();
	const part = await related(pkg, documentPart, 'connections', false);
	if (!part) return result;
	const root = await visioXml(pkg, part, 'DataConnections');
	for (const node of children(root, 'DataConnection')) {
		const id = attribute(node, 'ID') ?? '';
		const connectionString = attribute(node, 'ConnectionString') ?? '';
		const fileName = attribute(node, 'FileName');
		if (!/^\d{1,9}$/.test(id) || connectionString.length > 4096) continue;
		result.set(id, {
			id,
			connectionString,
			...(fileName ? { fileName: fileName.slice(0, 255) } : {}),
		});
	}
	return result;
}
