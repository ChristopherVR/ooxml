/** Parts A and B of the form: the client's details and the project brief. */
import type { Block } from '../../../../src/core/docx/index';
import { t, type Doc } from '../kit/builder';
import { pxFromCm } from '../kit/drawing';
import { table, tw } from '../kit/tables';
import { logoPlaceholder } from './art';
import { entry, formGrid, label, NO_BORDERS, partHeading, type Controls } from './controls';
import { INK, TEXT_WIDTH } from './styles';

export function clientDetails(doc: Doc, controls: Controls): Block[] {
	return [
		partHeading(doc, 'A', 'Client details'),
		formGrid(doc, [
			[
				label('Client name'),
				entry(controls.text('Client name', 'client.name', 'Legal name of the client')),
				label('Trading as'),
				entry(controls.text('Trading name', 'client.tradingName', 'If different')),
			],
			[
				label('Contact'),
				entry(controls.text('Contact person', 'client.contact', 'Full name')),
				label('Position'),
				entry(controls.text('Position', 'client.position', 'Role in the organisation')),
			],
			[
				label('Email'),
				entry(controls.text('Email', 'client.email', 'name@example.org')),
				label('Telephone'),
				entry(controls.text('Telephone', 'client.phone', 'Including the area code')),
			],
			[
				label('Address'),
				entry(
					controls.text(
						'Billing address',
						'client.address',
						'Street, town and postcode. Press Enter for a new line.',
						true,
					),
					{ span: 3 },
				),
			],
			[
				label('Client type'),
				entry(
					controls.choice(
						'dropDownList',
						'Client type',
						'client.type',
						['Sole trader', 'Partnership', 'Company', 'Charity or community group', 'Public body'],
						'Choose a type',
					),
				),
				label('Found us via'),
				entry(
					controls.choice(
						'comboBox',
						'Referral source',
						'client.referral',
						['Recommendation', 'Our portfolio', 'An event or talk', 'Web search'],
						'Choose, or type your own',
					),
				),
			],
			[
				label('Reach me by'),
				entry([
					...controls.check('Contact by email', 'client.byEmail', 'Email', true),
					...controls.check('Contact by telephone', 'client.byPhone', 'Phone'),
					...controls.check('Contact by post', 'client.byPost', 'Post'),
				]),
				label('Start date'),
				entry(controls.date('Preferred start date', 'project.start', 'Choose a date')),
			],
		]),
	];
}

const SERVICES: [string, string, boolean][] = [
	['Brand identity', 'services.identity', true],
	['Website design', 'services.web', false],
	['Editorial and print', 'services.print', true],
	['Signage and wayfinding', 'services.signage', false],
	['Illustration', 'services.illustration', false],
	['Art direction for photography', 'services.photography', false],
];

export function projectBrief(doc: Doc, controls: Controls): Block[] {
	const brief = controls.block('richText', 'Project brief', 'project.brief', {
		showingPlaceholder: true,
	});
	const gallery = controls.verbatim(
		'inline',
		'docPartList',
		'Standard scope wording',
		'project.scopeWording',
		'<w:docPartList><w:docPartGallery w:val="Quick Parts"/></w:docPartList>',
		'<w:showingPlcHdr/>',
	);
	const logo = doc.picture(logoPlaceholder(), {
		name: 'logo-placeholder',
		widthPx: pxFromCm(4.6),
		heightPx: pxFromCm(2.81),
		alt: 'Placeholder for the client logo.',
		title: 'Client logo',
	});
	const half = Math.floor(TEXT_WIDTH / 2);
	return [
		partHeading(doc, 'B', 'Project brief'),
		doc.p(
			'Hint',
			'Tell us what you need and why, in your own words. Formatting, lists and pictures are all welcome.',
			{ keepNext: true },
		),
		// Three paragraphs with the same shading and borders: Word draws them as one panel.
		...[
			'Describe the project: its purpose, its audience and what success looks like.',
			'',
			'',
			'',
			'',
			'',
		].map((text, index, all) =>
			doc.p('FieldValue', text ? t(text, { style: 'PlaceholderText' }) : '', {
				contentControl: brief,
				shadingFill: `#${INK.paper}`,
				borders: {
					bottom: { style: 'single', color: `#${INK.slate}`, spacePoints: 6 },
					top: { style: 'single', color: `#${INK.paper}`, spacePoints: 6 },
				},
				spacingAfterTwips: tw(index === all.length - 1 ? 240 : 0),
				indentLeftTwips: tw(113),
				indentRightTwips: tw(113),
			}),
		),
		table(doc, {
			style: 'FormGrid',
			grid: [half, TEXT_WIDTH - half],
			headerRows: 0,
			borders: NO_BORDERS,
			look: {
				firstRow: false,
				lastRow: false,
				firstColumn: false,
				lastColumn: false,
				noHBand: true,
				noVBand: true,
			},
			cantSplit: true,
			rows: [
				[
					{
						content: [
							doc.p('FieldLabel', 'Services requested', { spacingAfterTwips: tw(60) }),
							...SERVICES.map(([name, tag, checked]) =>
								doc.p('FieldValue', controls.check(name, tag, name, checked)),
							),
						],
					},
					{
						content: [
							doc.p('FieldLabel', 'Your current logo, if you have one', {
								spacingAfterTwips: tw(60),
							}),
							doc.p('FieldValue', [
								{ ...logo, contentControl: controls.picture('Client logo', 'client.logo') },
							]),
							doc.p('FieldLabel', 'Standard scope wording', {
								spacingBeforeTwips: tw(120),
								spacingAfterTwips: tw(40),
							}),
							doc.p('FieldValue', [
								{
									text: 'Choose a building block',
									style: 'PlaceholderText',
									contentControl: gallery,
								},
							]),
						],
					},
				],
			],
		}),
	];
}
