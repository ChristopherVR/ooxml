/**
 * The Aerie keynote's theme and slide master: a dark deck (the master maps the background to
 * the dark slots), a night backdrop and a quiet footer. The layouts are in `layouts.ts`.
 */
import type { TemplateSpec } from '../kit/template';
import { pictureBackgroundXml, rawShape } from '../kit/xml';
import { backdrop, duskSky } from './art-scenes';
import { layouts, number, title, titleStyle } from './layouts';
import { C, G, T } from './style';

export { LAYOUT } from './layouts';

export function template(): TemplateSpec {
	return {
		title: 'Aerie: the weather, from your own roof',
		subject: 'Nimbuswright launch keynote for the Aerie home weather station',
		theme: {
			name: 'Aerie Night',
			dark: true,
			colors: {
				dk1: C.night,
				lt1: C.text,
				dk2: C.panel,
				lt2: C.muted,
				accent1: C.amber,
				accent2: C.sky,
				accent3: C.violet,
				accent4: C.coral,
				accent5: '#5EE0B5',
				accent6: C.soft,
				hlink: C.sky,
				folHlink: '#B7A6FF',
			},
		},
		master: {
			background: pictureBackgroundXml('rIdNight'),
			images: { rIdNight: 'night.png' },
			shapes:
				rawShape({
					id: 20,
					name: 'Footer rule',
					x: G.left,
					y: 656,
					w: G.width,
					h: 1,
					fill: { solid: 'tx2', alpha: 0.22 },
				}) +
				rawShape({
					id: 21,
					name: 'Footer brand',
					x: G.left,
					y: 671,
					w: 500,
					h: 16,
					runs: [
						{ text: 'AERIE', size: 9, color: 'tx1', font: 'major', bold: true, tracking: 380 },
						{ text: '    BY NIMBUSWRIGHT', size: 9, color: 'tx2', font: 'major', tracking: 300 },
					],
				}),
			placeholders: [
				{ ...title(2), name: 'Title Placeholder' },
				{
					id: 3,
					name: 'Text Placeholder',
					type: 'body',
					idx: 1,
					x: G.left,
					y: G.content,
					w: G.width,
					h: 420,
					prompt: 'Text',
				},
				{ ...number(4), name: 'Slide Number Placeholder' },
			],
			titleStyle,
			bodyStyle: [
				{ size: T.lead, color: 'tx1', before: 14, leading: 118 },
				{
					size: T.body,
					color: 'tx2',
					bullet: '•',
					bulletColor: 'accent1',
					margin: 22,
					indent: -22,
					before: 6,
					leading: 118,
				},
				{ size: T.small, color: 'tx2', margin: 44, before: 4 },
			],
		},
		layouts,
		media: { 'night.png': backdrop(), 'dusk.png': duskSky() },
	};
}
