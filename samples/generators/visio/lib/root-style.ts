/** The root style sheet ("No Style") and the element every style sheet is written as. */
import { FONTS } from '../../lib/fiction';
import { attrs, cell, colorCell, pt, row, section } from './xml';

/** "No Style": the root every other style and shape inherits its defaults from. */
export function rootStyle(ink: string, themeCells: string): string {
	const cells = [
		cell('EnableLineProps', 1),
		cell('EnableFillProps', 1),
		cell('EnableTextProps', 1),
		cell('HideForApply', 0),
		cell('LineWeight', pt(0.75), undefined, 'PT'),
		colorCell('LineColor', ink),
		cell('LinePattern', 1),
		cell('Rounding', 0),
		cell('EndArrowSize', 2),
		cell('BeginArrow', 0),
		cell('EndArrow', 0),
		cell('LineCap', 0),
		cell('BeginArrowSize', 2),
		cell('LineColorTrans', 0),
		cell('CompoundType', 0),
		colorCell('FillForegnd', '#FFFFFF'),
		colorCell('FillBkgnd', '#FFFFFF'),
		cell('FillPattern', 1),
		colorCell('ShdwForegnd', ink),
		cell('ShdwPattern', 0),
		cell('FillForegndTrans', 0),
		cell('FillBkgndTrans', 0),
		cell('ShdwForegndTrans', 0),
		cell('ShapeShdwType', 0),
		cell('ShapeShdwOffsetX', 0),
		cell('ShapeShdwOffsetY', 0),
		cell('ShapeShdwObliqueAngle', 0),
		cell('ShapeShdwScaleFactor', 1),
		cell('ShapeShdwBlur', 0),
		cell('ShapeShdwShow', 0),
		cell('LeftMargin', 0),
		cell('RightMargin', 0),
		cell('TopMargin', 0),
		cell('BottomMargin', 0),
		cell('VerticalAlign', 1),
		cell('TextBkgnd', 0),
		cell('DefaultTabStop', 0.5),
		cell('TextDirection', 0),
		cell('TextBkgndTrans', 0),
		// 100 means "not themed": no quick style is applied until a user picks one.
		cell('QuickStyleLineColor', 100),
		cell('QuickStyleFillColor', 100),
		cell('QuickStyleShadowColor', 100),
		cell('QuickStyleFontColor', 100),
		cell('QuickStyleLineMatrix', 100),
		cell('QuickStyleFillMatrix', 100),
		cell('QuickStyleEffectsMatrix', 100),
		cell('QuickStyleFontMatrix', 100),
		cell('QuickStyleType', 0),
		cell('QuickStyleVariation', 0),
		cell('LineGradientEnabled', 0),
		cell('FillGradientEnabled', 0),
	];
	const character = row({ IX: 0 }, [
		cell('Font', FONTS.body),
		colorCell('Color', ink),
		cell('Style', 0),
		cell('Case', 0),
		cell('Pos', 0),
		cell('FontScale', 1),
		cell('Size', pt(9), undefined, 'PT'),
		cell('DblUnderline', 0),
		cell('Overline', 0),
		cell('Strikethru', 0),
		cell('DoubleStrikethrough', 0),
		cell('Letterspace', 0),
		cell('ColorTrans', 0),
		cell('AsianFont', 0),
		cell('ComplexScriptFont', 0),
		cell('ComplexScriptSize', -1),
	]);
	const paragraph = row({ IX: 0 }, [
		cell('IndFirst', 0),
		cell('IndLeft', 0),
		cell('IndRight', 0),
		cell('SpLine', -1.2),
		cell('SpBefore', 0),
		cell('SpAfter', 0),
		cell('HorzAlign', 1),
		cell('Bullet', 0),
		cell('BulletStr', ''),
		cell('BulletFont', 0),
		cell('BulletFontSize', -1),
		cell('TextPosAfterBullet', 0),
		cell('Flags', 0),
	]);
	return styleElement(0, 'No Style', undefined, [
		...cells,
		// Every shape resolves theme formulas with its own scheme cells; they inherit these.
		themeCells,
		section('Character', [character]),
		section('Paragraph', [paragraph]),
	]);
}

export function styleElement(
	id: number,
	name: string,
	parent: number | undefined,
	content: readonly string[],
): string {
	const names = { ID: id, NameU: name, IsCustomNameU: 1, Name: name, IsCustomName: 1 };
	const parents =
		parent === undefined ? {} : { LineStyle: parent, FillStyle: parent, TextStyle: parent };
	return `<StyleSheet${attrs({ ...names, ...parents })}>${content.join('')}</StyleSheet>`;
}
