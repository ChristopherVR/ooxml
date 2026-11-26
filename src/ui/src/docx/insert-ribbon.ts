import { EQUATION_GALLERY } from 'ooxml-core/docx/ui';
import { canCaptureDisplay } from '../form/display-capture';
import { SYMBOLS } from './insert-text-commands';
import { menuGallery } from './ribbon-gallery-menu';
import { group, menuSelect, stack, tool } from './ribbon-parts';

const PAGE_NUMBER_OPTIONS: Array<[string, string]> = [
	['top:left', 'Top left'],
	['top:center', 'Top center'],
	['top:right', 'Top right'],
	['bottom:left', 'Bottom left'],
	['bottom:center', 'Bottom center'],
	['bottom:right', 'Bottom right'],
	['top:left:pageOfTotal', 'Top left - Page X of Y'],
	['top:center:pageOfTotal', 'Top center - Page X of Y'],
	['top:right:pageOfTotal', 'Top right - Page X of Y'],
	['bottom:left:pageOfTotal', 'Bottom left - Page X of Y'],
	['bottom:center:pageOfTotal', 'Bottom center - Page X of Y'],
	['bottom:right:pageOfTotal', 'Bottom right - Page X of Y'],
];
export function buildInsertGroups(): HTMLElement[] {
	const screenshot = menuSelect(
		'Screenshot',
		'picture',
		[
			['full', 'Screenshot'],
			['clip', 'Screen Clipping'],
		],
		(value) => ({ type: 'insertScreenshot', clip: value === 'clip' }),
		{ momentary: true },
	);
	const captureSelect = screenshot.querySelector('select');
	if (captureSelect && !canCaptureDisplay()) {
		captureSelect.disabled = true;
		captureSelect.dataset.unavailable = '';
	}
	return [
		group(
			'Pages',
			tool('Cover page', 'coverPage', { type: 'coverPage' }, { large: true }),
			tool('Blank page', 'blankPage', { type: 'blankPage' }, { large: true }),
			tool(
				'Insert page break',
				'pageBreak',
				{ type: 'insertBreak', kind: 'page' },
				{ large: true, caption: 'Page break' },
			),
		),
		group(
			'Tables',
			tool('Insert table', 'table', { type: 'tablePicker' }, { large: true, caption: 'Table' }),
		),
		group(
			'Illustrations',
			tool(
				'Insert picture',
				'picture',
				{ type: 'insertPicture' },
				{ large: true, caption: 'Pictures', name: 'ribbon.name.insertPictures' },
			),
			screenshot,
			tool('Icons', 'insertIcon', { type: 'insertIcon' }, { large: true }),
			tool(
				'3D Models',
				'insertModel3d',
				{ type: 'insertLocalDrawing', kind: 'model3d' },
				{ large: true },
			),
			tool('Shapes', 'shapes', { type: 'shapePicker' }, { large: true }),
			stack(
				tool(
					'Insert chart',
					'chart',
					{ type: 'formatDialog', kind: 'insertChart' },
					{ inline: true, caption: 'Chart' },
				),
				tool(
					'Insert SmartArt',
					'smartArt',
					{ type: 'insertSmartArt' },
					{ inline: true, caption: 'SmartArt' },
				),
				tool(
					'Format picture',
					'formatPicture',
					{ type: 'formatPicture' },
					{ inline: true, caption: 'Format', name: 'ribbon.name.formatPicture' },
				),
			),
		),
		group(
			'Add-ins',
			tool('Get Add-ins', 'addIns', { type: 'officeAddIns', mode: 'getAddIns' }, { large: true }),
			tool('My Add-ins', 'addIns', { type: 'officeAddIns', mode: 'myAddIns' }, { large: true }),
		),
		group(
			'Media',
			tool('Online Video', 'picture', null, {
				large: true,
				unavailableReason: 'Inserting online video is not supported',
			}),
		),
		group(
			'Links',
			tool('Insert link', 'link', { type: 'link' }, { large: true, caption: 'Link' }),
			tool('Bookmark', 'bookmark', { type: 'formatDialog', kind: 'bookmark' }, { large: true }),
		),
		group(
			'Comments',
			tool(
				'Add comment',
				'comment',
				{ type: 'comments', key: 'add' },
				{
					large: true,
					caption: 'Comment',
				},
			),
		),
		group(
			'Header & Footer',
			menuSelect(
				'Header',
				'header',
				[
					['blank', 'Blank'],
					['first', 'Blank first page'],
					['even', 'Blank even pages'],
				],
				(value) => ({
					type: 'headerFooter',
					kind: 'header',
					slot: value === 'blank' ? 'default' : (value as 'first' | 'even'),
				}),
				{
					momentary: true,
				},
			),
			menuSelect(
				'Footer',
				'footer',
				[
					['blank', 'Blank'],
					['first', 'Blank first page'],
					['even', 'Blank even pages'],
				],
				(value) => ({
					type: 'headerFooter',
					kind: 'footer',
					slot: value === 'blank' ? 'default' : (value as 'first' | 'even'),
				}),
				{
					momentary: true,
				},
			),
			menuGallery(
				'Page number',
				'pageNumberIcon',
				'pageNumber',
				PAGE_NUMBER_OPTIONS,
				(value) => {
					const [position, align, style = 'plain'] = value.split(':') as [
						'top' | 'bottom',
						'left' | 'center' | 'right',
						('plain' | 'pageOfTotal')?,
					];
					return { type: 'pageNumber', position, align, style };
				},
				{ momentary: true },
			),
		),
		group(
			'Text',
			stack(
				tool(
					'Object',
					'insertObject',
					{ type: 'insertLocalDrawing', kind: 'object' },
					{ inline: true },
				),
				tool(
					'WordArt',
					'insertWordArt',
					{ type: 'insertWordArt' },
					{ inline: true, caption: 'WordArt' },
				),
				tool(
					'Text box',
					'textBox',
					{ type: 'formatDialog', kind: 'textBox' },
					{ inline: true, caption: 'Text Box' },
				),
			),
			stack(
				menuSelect(
					'Drop cap',
					'dropCap',
					[
						['none', 'None'],
						['drop', 'Dropped'],
						['margin', 'In margin'],
						['options', 'Drop Cap Options…'],
					],
					(value) =>
						value === 'options'
							? { type: 'formatDialog', kind: 'dropCap' }
							: { type: 'dropCap', value: value as 'none' },
					{ momentary: true, inline: true },
				),
				menuSelect(
					'Content control',
					'contentControl',
					[
						['richText', 'Rich text'],
						['text', 'Plain text'],
						['checkbox', 'Check box'],
						['dropDownList', 'Drop-down list'],
						['comboBox', 'Combo box'],
						['date', 'Date picker'],
						['properties', 'Properties…'],
					],
					(value) =>
						value === 'properties'
							? { type: 'formatDialog', kind: 'contentControl' }
							: { type: 'insertContentControl', value: value as 'richText' },
					{ momentary: true, inline: true },
				),
				menuSelect(
					'Date and time',
					'dateTime',
					[
						['long', 'long'],
						['short', 'short'],
						['time', 'time'],
						['datetime', 'datetime'],
					],
					(value) => ({ type: 'insertDateTime', value: value as 'long' }),
					{ momentary: true, inline: true },
				),
			),
		),
		group(
			'Symbols',
			menuSelect(
				'Equation',
				'equation',
				[
					['new', 'Insert New Equation'],
					['edit', 'Insert or Edit Equation…'],
					...EQUATION_GALLERY.map((entry): [string, string] => [entry.id, entry.name]),
				],
				(value) =>
					value === 'new'
						? { type: 'equation', key: 'new' }
						: value === 'edit'
							? { type: 'formatDialog', kind: 'equation' }
							: { type: 'insertEquation', id: value as 'area' },
				{ momentary: true },
			),
			menuSelect(
				'Symbol',
				'symbol',
				[...SYMBOLS.map((symbol): [string, string] => [symbol, symbol]), ['more', 'More Symbols…']],
				(value) =>
					value === 'more' ? { type: 'openSymbolGallery' } : { type: 'insertSymbol', value },
				{ momentary: true },
			),
		),
	];
}
