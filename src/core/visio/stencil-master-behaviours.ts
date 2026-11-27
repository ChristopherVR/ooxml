import { setCell } from './edit-geometry-cells';
import { attribute, children } from './sheet';
import type { VisioBuiltInMaster } from './stencil-masters';

/**
 * A control handle of a built-in master and the outline vertices it drives. Positions and
 * limits are those Visio 16's Basic Shapes (US units) masters report through the documented
 * object model (Controls.Row_1 X, Y, XCon, YCon and Prompt), placed on this package's own
 * outlines: the handle sits on the vertex it moves, X follows the drag and Y is locked
 * (YCon 1), as in Visio.
 */
interface ControlBehaviour {
	/** Controls.Row_1.X: the handle's default as a ratio of Width, inside Visio's BOUND limits. */
	x: (ratio: string) => string;
	ratio: number;
	prompt: string;
	/** Geometry1 rows (IX) whose X is the handle (`c`) or its mirror (`1-c`), as a ratio of Width. */
	rows: Readonly<Record<number, 'c' | '1-c'>>;
}

const CONTROLS: Readonly<Record<string, ControlBehaviour>> = {
	triangle: { x: (r) => `Width*${r}`, ratio: 0.5, prompt: 'Apex Position', rows: { 3: 'c' } },
	'right-triangle': {
		x: (r) => `Width*${r}`,
		ratio: 0,
		prompt: 'Apex Position',
		rows: { 3: 'c' },
	},
	parallelogram: {
		x: (r) => `BOUND(Width*${r},0,FALSE,0,Width)`,
		ratio: 0.25,
		prompt: 'Side Offset',
		rows: { 4: 'c', 2: '1-c' },
	},
	trapezoid: {
		x: (r) => `BOUND(Width*${r},0,FALSE,0,Width*0.5)`,
		ratio: 0.25,
		prompt: 'Side Offset',
		rows: { 4: 'c', 3: '1-c' },
	},
	chevron: {
		x: (r) => `BOUND(Width*${r},0,FALSE,0,Width)`,
		ratio: 0.75,
		prompt: 'Thickness',
		rows: { 2: 'c', 4: 'c', 6: '1-c' },
	},
};

/** Basic Flowchart Shapes masters that carry Visio's Set to Default Size action. */
const DEFAULT_SIZE_STENCILS = new Set(['basic-flowchart']);
const NO_DEFAULT_SIZE = new Set(['flowchart-on-page-reference', 'flowchart-off-page-reference']);

function section(shape: Element, name: string, before: Element | null): Element {
	const node = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Section');
	node.setAttribute('N', name);
	shape.insertBefore(node, before);
	return node;
}
function namedRow(owner: Element, name: string): Element {
	const row = owner.ownerDocument!.createElementNS(owner.namespaceURI, 'Row');
	row.setAttribute('N', name);
	owner.appendChild(row);
	return row;
}
function textCell(row: Element, name: string, value: string, formula?: string): void {
	const cell = row.ownerDocument!.createElementNS(row.namespaceURI, 'Cell');
	cell.setAttribute('N', name);
	cell.setAttribute('V', value);
	if (formula) cell.setAttribute('F', formula);
	row.appendChild(cell);
}

/**
 * Give a built-in master the ShapeSheet behaviours of Visio's own master of that name: the
 * control handle of Triangle, Right triangle, Parallelogram, Trapezoid and Chevron (its outline
 * follows the handle through Geometry formulas the formula engine recalculates), and on the Basic
 * Flowchart Shapes masters Visio's User.DefaultWidth/DefaultHeight cells and the Set to Default
 * Size action (SETF of Width and Height, hidden while the shape is at its default size). Visio's
 * Properties (RUNADDON) and Resize with Text (TEXTHEIGHT) actions are not written, since this
 * editor cannot run them, and these stencils have no smart tags.
 */
export function appendBuiltInBehaviours(shape: Element, builtIn: VisioBuiltInMaster): void {
	const { width, height } = builtIn.master.size;
	const before =
		children(shape, 'Section').find((node) => attribute(node, 'N') === 'Connection') ?? null;
	const control = CONTROLS[builtIn.master.id];
	if (control) {
		const controls = section(shape, 'Control', before);
		const row = namedRow(controls, 'Row_1');
		setCell(row, 'X', width * control.ratio, control.x(String(control.ratio)));
		setCell(row, 'Y', height, 'Height*1');
		setCell(row, 'XDyn', width * control.ratio, 'Controls.Row_1');
		setCell(row, 'YDyn', height, 'Controls.Row_1.Y');
		setCell(row, 'XCon', 0);
		setCell(row, 'YCon', 1);
		setCell(row, 'CanGlue', 0);
		textCell(row, 'Prompt', control.prompt);
		const geometry = children(shape, 'Section').find(
			(node) => attribute(node, 'N') === 'Geometry' && attribute(node, 'IX') === '0',
		);
		for (const geometryRow of children(geometry, 'Row')) {
			const role = control.rows[Number(attribute(geometryRow, 'IX'))];
			if (!role) continue;
			const value = role === 'c' ? control.ratio : 1 - control.ratio;
			setCell(
				geometryRow,
				'X',
				value,
				role === 'c' ? 'Controls.Row_1/Width' : '1-Controls.Row_1/Width',
			);
		}
	}
	if (DEFAULT_SIZE_STENCILS.has(builtIn.stencil.id) && !NO_DEFAULT_SIZE.has(builtIn.master.id)) {
		const actions = section(shape, 'Actions', before);
		const action = namedRow(actions, 'SetDefaultSize');
		textCell(action, 'Menu', 'Set to Default Si&ze');
		textCell(
			action,
			'Action',
			'0',
			'SETF(GetRef(Width),User.DefaultWidth)+SETF(GetRef(Height),User.DefaultHeight)',
		);
		setCell(action, 'Checked', 0);
		setCell(action, 'Disabled', 0);
		setCell(action, 'Invisible', 1, 'AND(Height=User.DefaultHeight,Width=User.DefaultWidth)');
		setCell(action, 'BeginGroup', 0);
		const user = section(shape, 'User', before);
		for (const [name, value] of [
			['DefaultWidth', width],
			['DefaultHeight', height],
		] as const) {
			const row = namedRow(user, name);
			const cell = row.ownerDocument!.createElementNS(row.namespaceURI, 'Cell');
			cell.setAttribute('N', 'Value');
			cell.setAttribute('V', String(value));
			cell.setAttribute('U', 'IN');
			row.appendChild(cell);
			textCell(row, 'Prompt', '');
		}
	}
}
