// Guide-driven DrawingML custom paths, evaluated through the shared geometry engine.
import { createBuiltinVariables } from '../geometry/guide-formula-api';
import { evaluateFormula, parseFormula } from '../geometry/guide-formula-eval';
import { NS, children, elements, first, type XmlElement } from './dom';
import type { DrawingPath, DrawingPathCommand } from './types';

const ARITY: Record<string, number> = {
	val: 1,
	abs: 1,
	sqrt: 1,
	'+-': 3,
	'*/': 3,
	'+/': 3,
	'?:': 3,
	at2: 2,
	cat2: 3,
	cos: 2,
	max: 2,
	min: 2,
	mod: 3,
	pin: 3,
	sat2: 3,
	sin: 2,
	tan: 2,
};

function operand(value: string | null, vars: Map<string, number>): number {
	if (value === null || !value.trim()) return NaN;
	const literal = Number(value);
	return Number.isFinite(literal) ? literal : (vars.get(value) ?? NaN);
}

/** Undefined means an unknown command, formula or unresolved coordinate: never invent a path. */
export function parseCustomGeometry(
	custom: XmlElement,
	width: number,
	height: number,
): DrawingPath[] | undefined {
	const pathList = first(custom, 'pathLst', NS.a);
	if (!pathList) return undefined;
	const pathNodes = children(pathList, 'path', NS.a);
	const firstPath = pathNodes[0];
	const contextWidth = Number(firstPath?.getAttribute('w')) || width;
	const contextHeight = Number(firstPath?.getAttribute('h')) || height;
	const vars = createBuiltinVariables({ w: contextWidth, h: contextHeight });
	for (const list of ['avLst', 'gdLst']) {
		const parent = first(custom, list, NS.a);
		if (!parent) continue;
		for (const guide of children(parent, 'gd', NS.a)) {
			const name = guide.getAttribute('name');
			const formula = parseFormula(guide.getAttribute('fmla') ?? '');
			if (
				!name ||
				ARITY[formula.op] !== formula.args.length ||
				formula.args.some((arg) => !Number.isFinite(operand(arg, vars)))
			)
				return undefined;
			const result = evaluateFormula(formula, vars);
			if (!Number.isFinite(result)) return undefined;
			vars.set(name, result);
		}
	}
	const paths: DrawingPath[] = [];
	for (const path of pathNodes) {
		const commands: DrawingPathCommand[] = [];
		const value = (node: XmlElement, name: string) => operand(node.getAttribute(name), vars);
		for (const node of elements(path)) {
			if (node.namespaceURI !== NS.a) return undefined;
			const points = children(node, 'pt', NS.a).map((point) => ({
				x: value(point, 'x'),
				y: value(point, 'y'),
			}));
			if (points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y)))
				return undefined;
			const [p0, p1, p2] = points;
			switch (node.localName) {
				case 'moveTo':
				case 'lnTo':
					if (points.length !== 1 || !p0) return undefined;
					commands.push({ op: node.localName === 'moveTo' ? 'M' : 'L', ...p0 });
					break;
				case 'quadBezTo':
					if (points.length !== 2 || !p0 || !p1) return undefined;
					commands.push({ op: 'Q', x1: p0.x, y1: p0.y, ...p1 });
					break;
				case 'cubicBezTo':
					if (points.length !== 3 || !p0 || !p1 || !p2) return undefined;
					commands.push({ op: 'C', x1: p0.x, y1: p0.y, x2: p1.x, y2: p1.y, ...p2 });
					break;
				case 'arcTo': {
					const wR = value(node, 'wR');
					const hR = value(node, 'hR');
					const start = value(node, 'stAng');
					const swing = value(node, 'swAng');
					if (![wR, hR, start, swing].every(Number.isFinite)) return undefined;
					commands.push({ op: 'A', wR, hR, startAngle: start / 60000, swingAngle: swing / 60000 });
					break;
				}
				case 'close':
					commands.push({ op: 'Z' });
					break;
				default:
					return undefined;
			}
		}
		const pathWidth = path.hasAttribute('w') ? Number(path.getAttribute('w')) : contextWidth;
		const pathHeight = path.hasAttribute('h') ? Number(path.getAttribute('h')) : contextHeight;
		if (![pathWidth, pathHeight].every((n) => Number.isFinite(n) && n > 0)) return undefined;
		paths.push({
			commands,
			width: pathWidth,
			height: pathHeight,
			...(path.getAttribute('fill') ? { fill: path.getAttribute('fill')! } : {}),
			...(['0', 'false'].includes(path.getAttribute('stroke') ?? '') ? { stroke: false } : {}),
		});
	}
	return paths.length ? paths : undefined;
}
