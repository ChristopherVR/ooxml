// Test-only: the element contract (src/ui/README.md, "How the elements are built") as checks. Each
// check looks at one element and returns why it fails, or undefined when it passes.
import { OfficeElement } from '../base';
import { CONTRACT_REVISION, defineElement } from '../registry';
import { COMMON_CONTROL_CSS } from '../styles';
import { OFFICE_TOKENS } from '../tokens';
import type { DefinedElement, ElementClass } from './elements';
import { leftOver, type ResourceTracker } from './resources';

export type CheckId =
	| 'office-element'
	| 'public-registration'
	| 'contract-stamp'
	| 'declared-properties'
	| 'attribute-properties'
	| 'early-properties'
	| 'token-styles'
	| 'disconnect-cleanup';

export interface Check {
	id: CheckId;
	/** What a passing element does, for the test title. */
	rule: string;
	/** Checks of the Lit contract skip elements that are not `OfficeElement`s. */
	officeOnly: boolean;
	failure(element: DefinedElement, context: CheckContext): Promise<string | undefined>;
}

export interface CheckContext {
	host: HTMLElement;
	resources: ResourceTracker;
}

const STAMP = Symbol.for('office-ui.web-control-contract');
let aliases = 0;

const wait = (ms: number) => new Promise<'waited'>((resolve) => setTimeout(resolve, ms, 'waited'));

/** Lets the update and the zero-delay timers of a mount run; fails when the update never ends. */
async function settle(element: Element): Promise<void> {
	const done = (element as Partial<OfficeElement>).updateComplete?.then(() => 'updated');
	if ((await Promise.race([done, wait(2000)])) === 'waited')
		throw new Error('updateComplete never settles');
	await wait(0);
}

/** The reactive properties a host can set: declared, not internal state. */
function publicProperties(ctor: ElementClass): [string, { type?: unknown }][] {
	return [...(ctor.elementProperties ?? [])]
		.filter(([name, options]) => typeof name === 'string' && !options.state)
		.map(([name, options]) => [name as string, options]);
}

const camel = (attribute: string): string =>
	attribute.replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase());

/** Whether instances of `ctor` have a settable `name`, from any class above the platform's. */
function hasSetter(ctor: ElementClass, name: string): boolean {
	for (
		let proto = ctor.prototype as object | null;
		proto && proto !== HTMLElement.prototype;
		proto = Object.getPrototypeOf(proto) as object | null
	)
		if (Object.getOwnPropertyDescriptor(proto, name)?.set) return true;
	return false;
}

/** A value for `name` that differs from its default, or undefined when it is not a scalar. */
function probeValue(element: Record<string, unknown>, name: string, type: unknown): unknown {
	const current = element[name];
	if (type === Boolean || typeof current === 'boolean') return !current;
	if (type === Number || typeof current === 'number')
		return Number.isFinite(current) ? (current as number) + 7 : 7;
	if (type === String || typeof current === 'string') return 'conformance-early-value';
	return undefined;
}

const unfallbacked = (css: string): string[] =>
	[...css.matchAll(/var\((--office-[\w-]+)\)/g)]
		.map((match) => match[1]!)
		.filter((name) => name in OFFICE_TOKENS);

export const CHECKS: readonly Check[] = [
	{
		id: 'office-element',
		rule: 'is an OfficeElement',
		officeOnly: false,
		failure: async ({ ctor }) =>
			ctor.prototype instanceof OfficeElement ? undefined : 'extends another base class',
	},
	{
		id: 'public-registration',
		rule: 'is defined by a public register function',
		officeOnly: false,
		failure: async (element) =>
			element.public ? undefined : 'defined on use, not by an exported register function',
	},
	{
		id: 'contract-stamp',
		rule: 'carries the contract stamp',
		officeOnly: false,
		failure: async ({ ctor }) =>
			(ctor as unknown as Record<symbol, unknown>)[STAMP] === CONTRACT_REVISION
				? undefined
				: 'not defined through the registry',
	},
	{
		id: 'declared-properties',
		rule: 'declares its reactive properties instead of shadowing them with fields',
		officeOnly: true,
		failure: async ({ tag, ctor }) => {
			const element = document.createElement(tag);
			const shadowed = publicProperties(ctor)
				.map(([name]) => name)
				.filter((name) => Object.hasOwn(element, name));
			return shadowed.length ? `own fields ${shadowed.join(', ')}` : undefined;
		},
	},
	{
		id: 'attribute-properties',
		rule: 'has a property for every attribute it reads',
		officeOnly: true,
		failure: async ({ ctor }) => {
			const missing = [...(ctor.watched ?? [])].filter((name) => !hasSetter(ctor, camel(name)));
			return missing.length ? `attribute only: ${missing.join(', ')}` : undefined;
		},
	},
	{
		id: 'early-properties',
		rule: 'keeps a property set before its definition loaded',
		officeOnly: true,
		failure: async ({ tag, ctor }, { host }) => {
			const reference = document.createElement(tag) as unknown as Record<string, unknown>;
			const probe = publicProperties(ctor)
				.map(([name, { type }]) => [name, probeValue(reference, name, type)] as const)
				.find(([, value]) => value !== undefined);
			if (!probe) return undefined;
			const [name, value] = probe;
			reference[name] = value;
			const expected = reference[name];
			const alias = `conformance-upgrade-${++aliases}`;
			const early = document.createElement(alias) as unknown as Record<string, unknown> & Element;
			early[name] = value;
			host.append(early);
			defineElement(customElements, alias, class extends ctor {});
			await settle(early);
			const actual = early[name];
			early.remove();
			return Object.is(actual, expected)
				? undefined
				: `${name} set before the upgrade reads ${String(actual)}, not ${String(expected)}`;
		},
	},
	{
		id: 'token-styles',
		rule: 'styles its shadow root from ?raw CSS through controlStyles, with token defaults',
		officeOnly: true,
		failure: async ({ ctor }) => {
			const sheets = (ctor.elementStyles ?? []).map((sheet) => sheet.cssText);
			if (!sheets.length) return 'no styles';
			if (sheets[0] !== COMMON_CONTROL_CSS) return 'styles not built with controlStyles';
			const bare = sheets.flatMap(unfallbacked);
			return bare.length ? `tokens without a default: ${[...new Set(bare)].join(', ')}` : undefined;
		},
	},
	{
		id: 'disconnect-cleanup',
		rule: 'lets go of window and document resources when removed',
		officeOnly: true,
		failure: async ({ tag }, { host, resources }) => {
			const before = resources.live();
			const element = document.createElement(tag);
			host.append(element);
			await settle(element);
			element.remove();
			await settle(element);
			const kept = leftOver(before, resources.live());
			return kept.length ? `still holds ${[...new Set(kept)].join(', ')}` : undefined;
		},
	},
];
