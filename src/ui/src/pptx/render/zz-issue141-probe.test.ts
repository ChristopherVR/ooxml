// @vitest-environment happy-dom
import { readFileSync, writeFileSync } from 'node:fs';
import { it } from 'vitest';
import { loadPresentation } from '../dom/load-presentation';
import { buildMorphTransitionPlan } from './morph-plan';

const DECK = 'C:/Users/Christopher/Downloads/Corporate-Deck-Abstract-Patterns (2).pptx';
const OUT = process.env.PROBE_OUT!;

it('probe', async () => {
	const bytes = readFileSync(DECK);
	const buf = new ArrayBuffer(bytes.byteLength);
	new Uint8Array(buf).set(bytes);
	const loaded = await loadPresentation(buf);
	const out: string[] = [];
	for (const [a, b] of [[6, 7], [17, 18]]) {
		const from = loaded.slides[a - 1]!;
		const to = loaded.slides[b - 1]!;
		out.push(`===== ${a} -> ${b}`);
		for (const s of [from, to]) for (const e of s.elements) out.push(`${s.slideNumber} ${e.id} ${(e as any).name} ${e.type} x=${Math.round(e.x)} y=${Math.round(e.y)} w=${Math.round(e.width)} h=${Math.round(e.height)} ${JSON.stringify(((e as any).text ?? '').slice(0, 40))} font=${(e as any).textSegments?.[0]?.style?.fontFamily} a=${(e as any).textSegments?.map((g: any) => g.style?.colorRef?.alpha).join(',')}`);
		const plan = buildMorphTransitionPlan(from, to, 2000, 'object');
		out.push(JSON.stringify(plan, (k, v) => (v instanceof Map ? Object.fromEntries(v) : v instanceof Set ? [...v] : k === 'rawXml' || k === 'textSegments' || k === 'paragraphs' ? undefined : v), 1));
	}
	writeFileSync(OUT, out.join('\n'));
	loaded.handler.dispose();
}, 120000);
