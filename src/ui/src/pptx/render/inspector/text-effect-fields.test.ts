import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import {
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from './inspector-field-model';
import { TEXT_EFFECT_SECTIONS } from './text-effect-fields';

const field = (section: string, id: string, ts: TextStyle | undefined, raw: string) =>
	inspectorFieldPatch(TEXT_EFFECT_SECTIONS, section, id, ts, raw);
const toggle = (section: string, ts: TextStyle | undefined, on: boolean) =>
	inspectorTogglePatch(TEXT_EFFECT_SECTIONS, section, ts, on);

describe('text effect sections', () => {
	it('offers text shadow, glow and reflection, each off for plain text', () => {
		expect(
			resolveInspectorSections(TEXT_EFFECT_SECTIONS, undefined).map((view) => [
				view.id,
				view.toggleable,
				view.on,
				view.toggleTestId,
				view.fields.length,
			]),
		).toStrictEqual([
			['shadow', true, false, 'text-fx-shadow-toggle', 0],
			['glow', true, false, 'text-fx-glow-toggle', 0],
			['reflection', true, false, 'text-fx-reflection-toggle', 0],
		]);
	});

	it('labels every section and field with a dictionary key', () => {
		const keys = TEXT_EFFECT_SECTIONS.flatMap((section) => [
			section.labelKey ?? '',
			...section.fields.map((f) => f.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});

	it('switches each effect on with the defaults the hand-built panels wrote', () => {
		expect(toggle('shadow', {}, true)).toStrictEqual({
			textShadowColor: '#000000',
			textShadowBlur: 4,
			textShadowOffsetX: 2,
			textShadowOffsetY: 2,
			textShadowOpacity: 0.5,
		});
		expect(toggle('glow', {}, true)).toStrictEqual({
			textGlowColor: '#ffff00',
			textGlowRadius: 6,
			textGlowOpacity: 0.6,
		});
		expect(toggle('reflection', {}, true)).toStrictEqual({
			textReflection: true,
			textReflectionBlur: 1,
			textReflectionStartOpacity: 0.5,
			textReflectionEndOpacity: 0,
			textReflectionOffset: 3,
		});
	});

	it('clears every field of an effect switched off', () => {
		expect(toggle('glow', { textGlowColor: '#ff0000' }, false)).toStrictEqual({
			textGlowColor: undefined,
			textGlowRadius: undefined,
			textGlowOpacity: undefined,
		});
		expect(toggle('shadow', { textShadowBlur: 3 }, false)).toStrictEqual({
			textShadowColor: undefined,
			textShadowBlur: undefined,
			textShadowOffsetX: undefined,
			textShadowOffsetY: undefined,
			textShadowOpacity: undefined,
		});
		expect(toggle('reflection', { textReflection: true }, false)).toStrictEqual({
			textReflection: undefined,
			textReflectionBlur: undefined,
			textReflectionStartOpacity: undefined,
			textReflectionEndOpacity: undefined,
			textReflectionOffset: undefined,
		});
	});

	it('treats a glow or shadow with only a size as on, as the renderer does', () => {
		const views = resolveInspectorSections(TEXT_EFFECT_SECTIONS, {
			textGlowRadius: 5,
			textShadowBlur: 2,
		});
		expect(views.map((view) => view.on)).toStrictEqual([true, true, false]);
	});

	it('shows the current values, falling back to the defaults, rounded for display', () => {
		const [shadow, glow, reflection] = resolveInspectorSections(TEXT_EFFECT_SECTIONS, {
			textShadowColor: 'FF0000',
			textShadowBlur: 3.6,
			textGlowColor: '#00ff00',
			textReflection: true,
			textReflectionBlur: 1.26,
		});
		expect(shadow!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['color', 'color', '#FF0000'],
			['opacity', 'number', 0.5],
			['blur', 'number', 4],
			['offsetX', 'number', 2],
			['offsetY', 'number', 2],
		]);
		expect(glow!.fields.map((f) => [f.id, f.value, f.min, f.max])).toStrictEqual([
			['color', '#00ff00', undefined, undefined],
			['opacity', 0.6, 0, 1],
			['size', 6, 1, 50],
		]);
		expect(reflection!.fields.map((f) => [f.id, f.value])).toStrictEqual([
			['blur', 1.3],
			['offset', 3],
			['startOpacity', 0.5],
			['endOpacity', 0],
		]);
	});

	it('writes one key per field, clamped to the offered range', () => {
		const glowing: TextStyle = { textGlowColor: '#ff0000' };
		expect(field('glow', 'color', glowing, '#123456')).toStrictEqual({ textGlowColor: '#123456' });
		expect(field('glow', 'size', glowing, '400')).toStrictEqual({ textGlowRadius: 50 });
		expect(field('glow', 'size', glowing, '0')).toStrictEqual({ textGlowRadius: 1 });
		expect(field('glow', 'opacity', glowing, '1.5')).toStrictEqual({ textGlowOpacity: 1 });
		expect(field('shadow', 'offsetX', {}, '-80')).toStrictEqual({ textShadowOffsetX: -50 });
		expect(field('reflection', 'endOpacity', {}, '0.25')).toStrictEqual({
			textReflectionEndOpacity: 0.25,
		});
		expect(field('glow', 'size', glowing, '')).toBeUndefined();
	});
});
