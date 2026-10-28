/**
 * signal-inputs.test-support.ts: declare the signal inputs and outputs of a
 * component tree for this package's plain JIT TestBed.
 *
 * The Angular compiler discovers `input()`, `model()` and `output()` fields
 * from the source; plain JIT (what vitest runs here, see `vitest.config.ts`)
 * only sees decorators, so a parent template binding `[slide]="..."` on a
 * child would overwrite the child's input signal with a raw value. The
 * Angular CLI's JIT mode fixes this with a source transform that adds
 * `@Input({ isSignal: true })` / `@Output()` to such fields; this helper adds
 * the same metadata at runtime, so a test can render a real component tree
 * (for example the thumbnail rail with real slide canvases inside).
 *
 * Fields are read from the authored sources under `src/viewer`, and the tree
 * is walked through each component's standalone `imports`. Call it before the
 * first `TestBed.configureTestingModule` that compiles these components.
 *
 * Not a `.test.ts` file on purpose: a shared helper is not a suite.
 */
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { Input, Output } from '@angular/core';
import type { Type } from '@angular/core';

interface SignalFields {
	inputs: { name: string; required: boolean }[];
	outputs: string[];
}

const FIELD = /readonly\s+(\w+)\s*=\s*(input|model|output)(\.required)?\s*[<(]/gu;

function sourceFiles(dir: string, out: string[] = []): string[] {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const file = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			sourceFiles(file, out);
		} else if (entry.name.endsWith('.ts') && !/\.(test|test-support)\.ts$/u.test(entry.name)) {
			out.push(file);
		}
	}
	return out;
}

/** Signal fields of every class declared under `dir`, by class name. */
function signalFieldsByClass(dir: string): Map<string, SignalFields> {
	const byClass = new Map<string, SignalFields>();
	for (const file of sourceFiles(dir)) {
		const source = readFileSync(file, 'utf8');
		const classes = [...source.matchAll(/export class (\w+)/gu)];
		classes.forEach((match, i) => {
			const body = source.slice(match.index, classes[i + 1]?.index ?? source.length);
			const fields: SignalFields = { inputs: [], outputs: [] };
			for (const [, name, kind, required] of body.matchAll(FIELD)) {
				if (kind === 'output') {
					fields.outputs.push(name!);
				} else {
					fields.inputs.push({ name: name!, required: Boolean(required) });
					if (kind === 'model') {
						fields.outputs.push(`${name!}Change`);
					}
				}
			}
			byClass.set(match[1]!, fields);
		});
	}
	return byClass;
}

const declared = new WeakSet<object>();

/** The standalone `imports` of a decorated component, flattened. */
function importsOf(type: Type<unknown>): unknown[] {
	const annotations = (type as unknown as { __annotations__?: { imports?: unknown[] }[] })
		.__annotations__;
	return (annotations ?? []).flatMap((annotation) => annotation.imports ?? []).flat(Infinity);
}

/**
 * Add JIT signal input/output metadata to `root` and every component it
 * imports (transitively). `dir` is the directory whose sources declare them.
 */
export function declareSignalFields(root: Type<unknown>, dir: string): void {
	const byClass = signalFieldsByClass(dir);
	const visit = (type: unknown): void => {
		if (typeof type !== 'function' || declared.has(type)) {
			return;
		}
		declared.add(type);
		const fields = byClass.get(type.name);
		for (const field of fields?.inputs ?? []) {
			Input({ isSignal: true, alias: field.name, required: field.required } as Input)(
				type.prototype,
				field.name,
			);
		}
		for (const name of fields?.outputs ?? []) {
			Output(name)(type.prototype, name);
		}
		importsOf(type as Type<unknown>).forEach(visit);
	};
	visit(root);
}
