import { afterEach, describe, expect, it } from 'vitest';
import {
	DEFAULT_BEHAVIOUR,
	MAX_UNDO_LEVELS,
	behaviourChanged,
	mergeBehaviour,
	readBehaviour,
	writeBehaviour,
} from './viewer-behaviour';

afterEach(() => localStorage.clear());

const TABS = ['home', 'insert', 'design', 'data', 'process', 'review', 'view', 'help'];

describe('behaviour settings', () => {
	it('starts with Office defaults: helpers on, every tab and toolbar command shown', () => {
		expect(DEFAULT_BEHAVIOUR).toMatchObject({
			miniToolbar: true,
			livePreview: true,
			autoSizeText: true,
			checkSpelling: true,
			displayLanguage: '',
			screenTips: 'descriptions',
			undoLevels: MAX_UNDO_LEVELS,
			'tab-home': true,
			'tab-help': true,
			'qat-save': true,
		});
	});

	it('takes values of the right type and ignores the rest', () => {
		const next = mergeBehaviour(DEFAULT_BEHAVIOUR, {
			miniToolbar: false,
			livePreview: 'yes',
			screenTips: 'loud',
			displayLanguage: 'klingon',
			undoLevels: '7',
			unknown: true,
		});
		expect(next.miniToolbar).toBe(false);
		expect(next.livePreview).toBe(true);
		expect(next.screenTips).toBe('descriptions');
		expect(next.displayLanguage).toBe('');
		expect(next.undoLevels).toBe(7);
		expect(Object.keys(next)).not.toContain('unknown');
	});

	it('clamps the undo depth and accepts only supported languages', () => {
		expect(mergeBehaviour(DEFAULT_BEHAVIOUR, { undoLevels: 500 }).undoLevels).toBe(MAX_UNDO_LEVELS);
		expect(mergeBehaviour(DEFAULT_BEHAVIOUR, { undoLevels: 0 }).undoLevels).toBe(1);
		expect(mergeBehaviour(DEFAULT_BEHAVIOUR, { displayLanguage: 'zh-CN' }).displayLanguage).toBe(
			'zh-CN',
		);
	});

	it('never leaves the ribbon without a tab', () => {
		const none = Object.fromEntries(TABS.map((id) => [`tab-${id}`, false]));
		const next = mergeBehaviour(DEFAULT_BEHAVIOUR, none);
		expect(next['tab-home']).toBe(true);
		expect(next['tab-insert']).toBe(false);
	});

	it('round-trips through storage and survives junk or missing storage', () => {
		writeBehaviour({ ...DEFAULT_BEHAVIOUR, checkSpelling: false, displayLanguage: 'es' });
		expect(readBehaviour()).toMatchObject({ checkSpelling: false, displayLanguage: 'es' });
		localStorage.setItem('ooxml-visio-options', '{not json');
		expect(readBehaviour()).toEqual(DEFAULT_BEHAVIOUR);
		localStorage.setItem('ooxml-visio-options', '"a string"');
		expect(readBehaviour()).toEqual(DEFAULT_BEHAVIOUR);
	});

	it('reports whether anything differs', () => {
		expect(behaviourChanged(DEFAULT_BEHAVIOUR, { ...DEFAULT_BEHAVIOUR })).toBe(false);
		expect(behaviourChanged(DEFAULT_BEHAVIOUR, { ...DEFAULT_BEHAVIOUR, 'qat-undo': false })).toBe(
			true,
		);
	});
});
