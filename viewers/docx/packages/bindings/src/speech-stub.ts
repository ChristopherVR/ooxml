import { vi } from 'vitest';

/** Test helper: a fake browser voice. `spoken` collects the texts; call `finish()` to end the last. */
export function stubSpeech() {
	const spoken: string[] = [];
	const ends: (() => void)[] = [];
	vi.stubGlobal('speechSynthesis', { speak: vi.fn(), cancel: vi.fn() });
	const synthesis = speechSynthesis as unknown as { speak: (utterance: { text: string }) => void };
	synthesis.speak = (utterance) => spoken.push(utterance.text);
	vi.stubGlobal(
		'SpeechSynthesisUtterance',
		class {
			constructor(public text: string) {}
			addEventListener(name: string, fn: () => void) {
				if (name === 'end') ends.push(fn);
			}
		},
	);
	return { spoken, finish: () => ends.at(-1)?.() };
}
