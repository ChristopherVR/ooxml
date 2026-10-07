import { describe, expect, it } from 'vitest';
import { transformCase } from './case-transform';

describe('transformCase', () => {
	it('handles every Word mode', () => {
		expect(transformCase('hello WORLD. this is it! ok', 'sentence')).toBe(
			'Hello world. This is it! Ok',
		);
		expect(transformCase('hello world', 'lower')).toBe('hello world');
		expect(transformCase('Hello world', 'upper')).toBe('HELLO WORLD');
		expect(transformCase("it's a dog-eat-dog world", 'title')).toBe("It's A Dog-Eat-Dog World");
		expect(transformCase('Hello World', 'toggle')).toBe('hELLO wORLD');
	});

	it('leaves a mid-sentence fragment lowercase when it does not start a sentence', () => {
		expect(transformCase('and more', 'sentence', false)).toBe('and more');
	});
});
