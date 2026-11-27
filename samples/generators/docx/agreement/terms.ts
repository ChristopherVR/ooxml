/** Parts D and E: the terms, in numbered clauses inside a locked group, and the signatures. */
import type { Block, ContentControl, Paragraph } from '../../../../src/core/docx/index';
import { rich, t, type Doc, type Inline } from '../kit/builder';
import { field, ref } from '../kit/fields';
import { line, NO_LINE, table, tw, type CellSpec } from '../kit/tables';
import { entry, formGrid, label, type Controls } from './controls';
import { INK, TEXT_WIDTH } from './styles';

const CO_OP = 'Quillhaven Design Co-operative';

/**
 * The clauses. The paragraph styles carry the numbering (1., 1.1, (a)), so the text holds no
 * numbers; the two references to other clauses are REF fields on bookmarks.
 */
export function terms(doc: Doc, controls: Controls): Paragraph[] {
	// A group the reader cannot edit or delete: the wording is fixed, the form around it is not.
	const group: ContentControl = controls.block('group', 'Terms and conditions', 'terms', {
		lock: 'sdtContentLocked',
	});
	const clause = (title: string, mark?: string) =>
		doc.p('Clause', title, { contentControl: group, ...(mark && { bookmarks: [mark] }) });
	const term = (content: Inline, mark?: string) =>
		doc.p('Term', content, { contentControl: group, ...(mark && { bookmarks: [mark] }) });
	const item = (text: string) => doc.p('TermItem', text, { contentControl: group });
	return [
		clause('Definitions and scope'),
		term(`In this agreement “we” means ${CO_OP} and “you” means the client named in Part A.`),
		term(
			rich(
				'The services are those described in the project brief (Part B) and priced in the schedule of fees (Part C). Anything else is a variation under clause {clause}.',
				{ clause: ref('_Ref_clause_variations', '4', '\\r \\h') },
			),
		),
		clause('Our commitments'),
		term(
			'We will carry out the services with reasonable skill and care, and by the dates agreed in the project plan.',
		),
		term(
			'We will name one member of the co-operative as your account manager, with authority to act for us.',
		),
		clause('Your commitments'),
		term('You will:'),
		item(
			'give us the content, access and decisions we reasonably need within five working days of a request;',
		),
		item('check drafts carefully, since we rely on your approval; and'),
		item('tell us promptly if anything in Part A changes.'),
		clause('Variations', '_Ref_clause_variations'),
		term(
			'Either of us may propose a change to the services. No change takes effect until both of us have agreed it in writing, including its effect on fees and dates.',
		),
		clause('Fees and payment'),
		term('Fees are as set out in Part C and exclude taxes.'),
		term(
			'We invoice the deposit on signature and the balance as each stage is approved. Invoices are payable within the payment terms chosen in Part C.',
		),
		term('If an invoice is more than 30 days overdue we may pause work until it is paid.'),
		clause('Intellectual property'),
		term('On payment in full, the rights in the final approved designs pass to you.'),
		term('We keep the rights in:'),
		item('preliminary concepts that you did not choose;'),
		item('our working files, tools and methods; and'),
		item(
			'fonts, photographs and other materials licensed from others, which stay subject to their licences.',
		),
		term('We may show the finished work in our portfolio unless you ask us in writing not to.'),
		clause('Confidentiality'),
		term(
			'Each of us will keep the other’s confidential information secret, and use it only for this agreement, for three years after it ends.',
		),
		clause('Liability'),
		term(
			'Nothing in this agreement limits liability that cannot lawfully be limited.',
			'_Ref_term_unlimited',
		),
		term(
			rich(
				'Subject to clause {clause}, our total liability under this agreement is limited to the fees paid.',
				{
					clause: ref('_Ref_term_unlimited', '8.1', '\\r \\h'),
				},
			),
		),
		clause('Ending the agreement'),
		term(
			'Either of us may end this agreement by giving 14 days’ written notice. You will pay for the work done up to the date it ends.',
		),
		clause('General'),
		term('This agreement is the whole agreement between us about the services.'),
		term(
			'It is governed by the law of the Northwind Coast, and the courts of Port Averly have jurisdiction.',
		),
	];
}

export function signatures(doc: Doc, controls: Controls): Block[] {
	const half = Math.floor((TEXT_WIDTH - 400) / 2);
	const party = (heading: string, key: string): CellSpec => ({
		content: [
			doc.p('FieldLabel', heading, { spacingAfterTwips: tw(520) }),
			doc.p('FieldValue', '', {
				borders: { top: { style: 'single', color: `#${INK.text}`, spacePoints: 3 } },
				spacingAfterTwips: tw(0),
			}),
			doc.p('FieldLabel', 'Signature', { spacingAfterTwips: tw(140) }),
			doc.p('FieldValue', [
				t('Name  ', { bold: true, fontSize: 8, color: `#${INK.slate}` }),
				...controls.text(`${key} signatory`, `sign.${key}.name`, 'Full name'),
			]),
			doc.p('FieldValue', [
				t('Position  ', { bold: true, fontSize: 8, color: `#${INK.slate}` }),
				...controls.text(`${key} position`, `sign.${key}.position`, 'Role'),
			]),
			doc.p('FieldValue', [
				t('Date  ', { bold: true, fontSize: 8, color: `#${INK.slate}` }),
				...controls.date(`${key} date`, `sign.${key}.date`, 'Choose a date'),
			]),
		],
		fill: `#${INK.paper}`,
		margins: { top: tw(170), bottom: tw(170), left: tw(200), right: tw(200) },
		borders: { top: line(INK.gold, 1.5) },
	});
	const none = {
		top: NO_LINE,
		left: NO_LINE,
		bottom: NO_LINE,
		right: NO_LINE,
		insideH: NO_LINE,
		insideV: NO_LINE,
	};
	return [
		doc.p('Heading1', [t('Part E', { style: 'PartLabel' }), t('   Signatures')]),
		doc.p(
			'Normal',
			[
				{
					text: '',
					fieldChar: 'begin',
					formField: {
						kind: 'checkbox',
						name: 'Authority',
						enabled: true,
						calcOnExit: false,
						checkBox: { default: false, sizeAuto: true },
					},
				},
				{ text: '', fieldCode: ' FORMCHECKBOX ' },
				{ text: '', fieldChar: 'end' },
				t(
					'  I confirm that I have read the terms in Part D and have authority to sign for the client.',
				),
			],
			{ keepNext: true, spacingAfterTwips: tw(160) },
		),
		table(doc, {
			style: 'FormGrid',
			grid: [half, 400, TEXT_WIDTH - 400 - half],
			headerRows: 0,
			borders: none,
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
					party('Signed for the client', 'Client'),
					{ content: '' },
					party(`Signed for ${CO_OP}`, 'Co-operative'),
				],
			],
		}),
		doc.p(
			'Hint',
			[
				t('Our standard terms, and what each clause means in practice, are explained at '),
				...field(
					'HYPERLINK "https://quillhaven.example/terms"',
					t('quillhaven.example/terms', { style: 'Hyperlink' }),
				),
				t('.'),
			],
			{ spacingBeforeTwips: tw(200) },
		),
		doc.p('Heading1', 'What happens next'),
		doc.p('Step', 'We countersign the agreement and send you a copy within two working days.'),
		doc.p('Step', 'Your account manager books a kick-off meeting and shares the project plan.'),
		doc.p('Step', 'We invoice the deposit. Work starts on the agreed date once it is paid.', {
			spacingAfterTwips: tw(280),
		}),
		doc.p('FieldLabel', 'For office use', { spacingAfterTwips: tw(80), keepNext: true }),
		formGrid(doc, [
			[
				label('Received by'),
				entry(
					controls.choice(
						'dropDownList',
						'Received by',
						'office.receivedBy',
						['Amara Diallo', 'Felix Brandauer', 'Zofia Marchetti', 'Rafael Monteverde'],
						'Choose a member',
					),
				),
				label('Received on'),
				entry(controls.date('Date received', 'office.received', 'Choose a date')),
			],
			[
				label('Invoice no.'),
				entry(controls.text('Deposit invoice number', 'office.invoice', 'Invoice number')),
				label('Checks'),
				entry([
					...controls.check('Credit check done', 'office.credit', 'Credit'),
					...controls.check('Conflict check done', 'office.conflict', 'Conflicts'),
				]),
			],
		]),
	];
}
