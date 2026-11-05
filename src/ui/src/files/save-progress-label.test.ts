import { describe, expect, it } from 'vitest';
import { saveProgressLabel, saveProgressStepChanged } from './save-progress-label';

describe('saveProgressLabel', () => {
	it('describes every phase in English', () => {
		expect(saveProgressLabel({ phase: 'uploading', loaded: 45, total: 100 })).toBe('Saving… 45%');
		expect(saveProgressLabel({ phase: 'checking' })).toBe('Checking for changes…');
		expect(saveProgressLabel({ phase: 'downloading', loaded: 3, total: 10 })).toBe(
			'Getting the latest version… 30%',
		);
		expect(saveProgressLabel({ phase: 'downloading', loaded: 3 })).toBe(
			'Getting the latest version…',
		);
		expect(saveProgressLabel({ phase: 'merging' })).toBe('Combining edits…');
	});

	it('translates before filling in the percentage, and clamps it', () => {
		const t = (text: string) => (text === 'Saving… {percent}%' ? 'Speichern… {percent} %' : text);
		expect(saveProgressLabel({ phase: 'uploading', loaded: 150, total: 100 }, t)).toBe(
			'Speichern… 100 %',
		);
		expect(saveProgressLabel({ phase: 'uploading', loaded: 0, total: 0 }, t)).toBe(
			'Speichern… 0 %',
		);
	});
});

describe('saveProgressStepChanged', () => {
	it('announces a new phase or each further 10 percent, not every fragment', () => {
		const at = (loaded: number) => ({ phase: 'uploading' as const, loaded, total: 100 });
		expect(saveProgressStepChanged(undefined, at(0))).toBe(true);
		expect(saveProgressStepChanged(at(11), at(15))).toBe(false);
		expect(saveProgressStepChanged(at(15), at(21))).toBe(true);
		expect(saveProgressStepChanged(at(100), { phase: 'checking' })).toBe(true);
		expect(saveProgressStepChanged({ phase: 'merging' }, { phase: 'merging' })).toBe(false);
	});
});
