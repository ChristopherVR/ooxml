import { readdirSync, readFileSync } from 'node:fs';
import { join, relative, resolve, sep } from 'node:path';
import { OFFICE_UI_TAGS } from './index.js';

/**
 * Builds the Custom Elements Manifest of this package, the machine-readable list of every element,
 * its attributes, properties and events that framework bindings and IDE tooling are generated from.
 * Properties and attributes come from the live classes (so they are exact); the description is the
 * class's doc comment and the events are the names the source fires (read from the source, so they
 * are only as complete as what is written as a string literal). Test and tooling use only.
 */
interface ElementClass extends CustomElementConstructor {
	elementProperties?: Map<
		string,
		{ type?: { name?: string }; attribute?: string | boolean; state?: boolean }
	>;
	observedAttributes?: string[];
	watched?: readonly string[];
}

/** Tests run from the package root, so the source and the manifest are found from there. */
const SRC = resolve(process.cwd(), 'src');
/** Where the prototype walk stops: members of these classes are not an element's own API. */
const BASES = new Set(['OfficeElement', 'LitElement', 'ReactiveElement', 'HTMLElement']);
const FIRE = /\b(?:fire|emit)\(\s*(?:this,\s*)?['"]([a-z][\w-]*-[\w-]+)['"]/g;
const EVENT_FIELD = /\b\w*[eE]vent\s*=\s*['"]([a-z][\w-]*-[\w-]+)['"]/g;

function sources(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) return sources(path);
		return /\.ts$/.test(entry.name) && !/\.(test|d)\.ts$/.test(entry.name) ? [path] : [];
	});
}

/** Doc comment directly above `export class <name>`, collapsed to plain text. */
function describeClass(text: string, name: string): string {
	const at = text.search(new RegExp(`export class ${name}\\b`));
	if (at < 0) return '';
	const before = text.slice(0, at);
	const end = before.lastIndexOf('*/');
	if (end < 0 || before.slice(end + 2).trim() !== '') return '';
	const start = before.lastIndexOf('/**', end);
	return before
		.slice(start + 3, end)
		.split('\n')
		.map((line) => line.replace(/^\s*\* ?/, ''))
		.join(' ')
		.replace(/\s+/g, ' ')
		.trim();
}

export function buildManifest(root: string = SRC) {
	const files = sources(SRC).map((path) => ({ path, text: readFileSync(path, 'utf8') }));
	const modules = [];
	for (const tag of [...OFFICE_UI_TAGS].sort()) {
		// The formatter may wrap the call, so match across whitespace.
		const definition = new RegExp(`definer\\(\\s*'${tag}',\\s*\\(\\)\\s*=>\\s*(\\w+)`);
		const file = files.find(({ text }) => definition.test(text));
		const className = file?.text.match(definition)?.[1];
		const ctor = customElements.get(tag) as ElementClass | undefined;
		if (!file || !className || !ctor) throw new Error(`${tag} has no class or definition`);
		const events = new Set<string>();
		for (const pattern of [FIRE, EVENT_FIELD])
			for (const match of file.text.matchAll(pattern)) events.add(match[1]!);
		const attributes = new Set<string>([
			...(ctor.observedAttributes ?? []),
			...(ctor.watched ?? []),
		]);
		const declared = [...(ctor.elementProperties ?? new Map())]
			// `state: true` properties are internal render state, not part of the element's API.
			.filter(([name, options]) => !name.startsWith('_') && !options.state)
			.map(([name, options]) => ({
				kind: 'field',
				name,
				...(options.type?.name ? { type: { text: options.type.name.toLowerCase() } } : {}),
				...(typeof options.attribute === 'string' ? { attribute: options.attribute } : {}),
			}));
		// Properties written as accessors (a controlled element's `state`, a clamped `value`) are
		// not in Lit's table, so read them off the prototype chain, up to the shared base class.
		const known = new Set(declared.map((member) => member.name));
		const accessors: Array<{ kind: string; name: string }> = [];
		for (
			let proto = ctor.prototype as object;
			proto && !BASES.has((proto.constructor as { name: string }).name);
			proto = Object.getPrototypeOf(proto) as object
		)
			for (const [name, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(proto)))
				if (descriptor.set && !name.startsWith('_') && !known.has(name)) {
					known.add(name);
					accessors.push({ kind: 'field', name });
				}
		const members = [...declared, ...accessors].sort((a, b) => a.name.localeCompare(b.name));
		const path = relative(root, file.path).replaceAll(sep, '/');
		modules.push({
			kind: 'javascript-module',
			path,
			declarations: [
				{
					kind: 'class',
					name: className,
					tagName: tag,
					customElement: true,
					description: describeClass(file.text, className),
					attributes: [...attributes].sort().map((name) => ({ name })),
					members,
					events: [...events].sort().map((name) => ({ name })),
				},
			],
			exports: [{ kind: 'custom-element-definition', name: tag, declaration: { name: className } }],
		});
	}
	return { schemaVersion: '2.1.0', readme: 'README.md', modules };
}
