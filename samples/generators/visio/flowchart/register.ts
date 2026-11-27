/** The process register: the table the flowchart's steps are linked to. */
import { RUN } from '../lib/palette';
import type { RecordsetSpec } from '../lib/recordset';
import type { Rich } from '../lib/text';

/** The process register: one row per step, linked to the step's shape. */
export const REGISTER = [
	['S-01', 'Validate order', 'Sales desk', 'Storefront', '0.1', '0.35', 'In place'],
	['S-02', 'Authorise payment', 'Finance', 'Payments ledger', '0.1', '0.22', 'In place'],
	['S-03', 'Payment authorised?', 'Sales desk', 'Storefront', '0', '0', 'In place'],
	['S-04', 'Check stock', 'Warehouse', 'Stock ledger', '0.2', '0.18', 'In place'],
	['S-05', 'In stock?', 'Warehouse', 'Stock ledger', '0', '0', 'In place'],
	['S-06', 'Raise back-order', 'Sales desk', 'Purchasing', '0.5', '1.4', 'Under review'],
	['S-07', 'Pick and pack', 'Warehouse', 'Pick lists', '0.6', '2.85', 'In place'],
	['S-08', 'Collect parcel', 'Carrier', 'Dispatch board', '4', '5.1', 'In place'],
	['S-09', 'Delivered?', 'Carrier', 'Parcel tracking', '0', '0', 'In place'],
	['S-10', 'Capture payment', 'Finance', 'Payments ledger', '0.1', '0.3', 'Planned'],
] as const;

export const REGISTER_RECORDSET: RecordsetSpec = {
	name: 'Process register',
	columns: [
		{ name: 'Step', label: 'Step', type: 'string' },
		{ name: 'Activity', label: 'Activity', type: 'string' },
		{ name: 'Owner', label: 'Owner', type: 'string' },
		{ name: 'System', label: 'System', type: 'string' },
		{ name: 'Duration', label: 'Duration (hours)', type: 'number' },
		{ name: 'Cost', label: 'Cost per order', type: 'number' },
		{ name: 'Status', label: 'Status', type: 'string' },
	],
	rows: REGISTER,
	primaryKey: 'Step',
};

/** A registered step: its name, and under it the system it runs in, read from its shape data. */
export function registered(step: (typeof REGISTER)[number][0], withSystem = true): Rich {
	const row = REGISTER.find((entry) => entry[0] === step);
	if (!row) throw new Error(`Unknown step ${step}`);
	if (!withSystem) return row[1];
	return [
		{ text: `${row[1]}\n`, style: RUN.step },
		{ text: row[3], field: 'Prop.System', style: RUN.small },
	];
}
