import { describe, expect, it } from 'vitest';
import { optionStrings } from 'ooxml-core/i18n';
import {
	optionText,
	optionsLabels,
	personalizeSection,
	unavailableCategory,
	unavailableControl,
	userInterfaceSection,
} from '../controls';

const english = { product: 'Word', t: (text: string) => text };
const french = {
	product: 'Word',
	t: (text: string) => optionStrings.fr[text] ?? text,
};

describe('shared Options categories', () => {
	it('builds an honest unavailable category with the product name filled in', () => {
		const category = unavailableCategory(english, 'proofing');
		expect(category).toMatchObject({
			id: 'proofing',
			label: 'Proofing',
			description: 'Change how Word corrects and formats your text.',
			disabled: 'Not available yet. Needs rich text formatting.',
			sections: [],
		});
		expect(
			unavailableCategory(english, 'save', { description: 'Customize how drawings are saved.' }),
		).toMatchObject({ description: 'Customize how drawings are saved.' });
	});

	it('uses the ids the products and the Office order rely on', () => {
		const ids = (
			[
				'proofing',
				'save',
				'language',
				'accessibility',
				'advanced',
				'ribbon',
				'qat',
				'addins',
				'trust',
			] as const
		).map((kind) => unavailableCategory(english, kind).id);
		expect(ids).toEqual([
			'proofing',
			'save',
			'language',
			'accessibility',
			'advanced',
			'ribbon',
			'qat',
			'add-ins',
			'trust',
		]);
	});

	it('translates labels, reasons and buttons through the product dictionary', () => {
		expect(unavailableCategory(french, 'addins').label).toBe('Compléments');
		expect(optionText(french, 'Make {product} easier to work with.')).toBe(
			'Facilitez l’utilisation de Word.',
		);
		expect(optionsLabels(french)).toMatchObject({
			cancel: 'Annuler',
			reset: 'Réinitialiser toutes les options',
		});
		expect(
			unavailableControl(
				french,
				{ kind: 'toggle', key: 'x', label: 'X' },
				'Not available yet. The editing behaviour is fixed.',
			).disabled,
		).toBe('Pas encore disponible. Le comportement de modification est fixe.');
	});

	it('lists the three user interface settings as disabled and keeps the personalize controls', () => {
		const ui = userInterfaceSection(english);
		expect(ui.controls.map((control) => [control.key, Boolean(control.disabled)])).toEqual([
			['miniToolbar', true],
			['livePreview', true],
			['screenTips', true],
		]);
		const personal = personalizeSection(
			english,
			[{ kind: 'text', key: 'author', label: 'User name' }],
			'Note',
		);
		expect(personal).toMatchObject({ id: 'personalize', description: 'Note' });
		expect(personal.controls).toHaveLength(1);
	});

	it('defines every shared string in every locale with the same placeholders', () => {
		const placeholders = (text: string) =>
			[...text.matchAll(/\{\w+\}/g)].map((match) => match[0]).sort();
		for (const [locale, strings] of Object.entries(optionStrings)) {
			expect(Object.keys(strings).sort(), locale).toEqual(Object.keys(optionStrings.en).sort());
			for (const [key, value] of Object.entries(strings)) {
				expect(value.trim(), `${locale}:${key}`).not.toBe('');
				expect(placeholders(value), `${locale}:${key}`).toEqual(placeholders(key));
			}
		}
	});
});
