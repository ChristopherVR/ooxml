import { cell, fixture, rectangle, relations, relation, shape, xml } from './test-fixtures.js';

export const drawingNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const visioThemeNamespace = 'http://schemas.microsoft.com/office/visio/2012/theme';
export const themeRelationship =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const extension = (value: string) => `<a:extLst><a:ext uri="generated">${value}</a:ext></a:extLst>`;
const id = (index: number) => `<v:schemeID schemeEnum="${index}"/>`;
export const solid = (color: string) => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
export const placeholder = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
/** Wholly generated test data, not a copy of any third-party theme or document. */
export function generatedTheme(
	options: { colorId?: number; effectId?: number; fill?: string; extraColors?: string } = {},
): string {
	const colors = ['102030', 'F0E0D0', '214365', '436587', '6587A9', '87A9CB', 'A9CBED', 'CBED0F'];
	const names = ['dk1', 'lt1', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
	const variants = ['123456', '345678', '56789A', '789ABC']
		.map(
			(value) =>
				`<v:variationClrScheme>${Array.from({ length: 7 }, (_, i) => `<v:varColor${i + 1}><a:srgbClr val="${i === 0 ? value : colors[i]}"/></v:varColor${i + 1}>`).join('')}</v:variationClrScheme>`,
		)
		.join('');
	const colorExtension = `<v:schemeID schemeEnum="${options.colorId ?? 42}"/><v:bkgnd><a:srgbClr val="FFEEDD"/></v:bkgnd><v:variationClrSchemeLst>${variants}</v:variationClrSchemeLst>`;
	const styleVariants = Array.from(
		{ length: 4 },
		() =>
			`<v:variationStyleScheme><v:varStyle fillIdx="4" lineIdx="4" fontIdx="4"/></v:variationStyleScheme>`,
	).join('');
	const styles = `${solid('FFFFFF')}${placeholder}${placeholder}${options.fill ?? placeholder}${placeholder}${placeholder}`;
	return `<a:theme xmlns:a="${drawingNamespace}" xmlns:v="${visioThemeNamespace}"><a:themeElements><a:clrScheme name="Generated">${names.map((name, i) => `<a:${name}><a:srgbClr val="${colors[i]}"/></a:${name}>`).join('')}${options.extraColors ?? ''}${extension(colorExtension)}</a:clrScheme><a:fmtScheme name="Generated"><a:fillStyleLst>${styles}</a:fillStyleLst><a:lnStyleLst>${Array.from({ length: 6 }, () => `<a:ln>${placeholder}</a:ln>`).join('')}</a:lnStyleLst></a:fmtScheme>${extension(`<v:fmtSchemeEx>${id(options.effectId ?? 42)}</v:fmtSchemeEx><v:fmtConnectorSchemeEx>${id(options.effectId ?? 42)}</v:fmtConnectorSchemeEx><v:fmtConnectorScheme><a:fillStyleLst>${styles}</a:fillStyleLst><a:lnStyleLst>${Array.from({ length: 6 }, () => `<a:ln>${placeholder}</a:ln>`).join('')}</a:lnStyleLst></v:fmtConnectorScheme><v:fontStylesGroup><v:fontStyles>${Array.from({ length: 6 }, () => '<v:fontProps><v:color><a:schemeClr val="phClr"/></v:color></v:fontProps>').join('')}</v:fontStyles></v:fontStylesGroup><v:variationStyleSchemeLst>${styleVariants}</v:variationStyleSchemeLst>`)}</a:themeElements></a:theme>`;
}
export const themeSelectors = (scheme = 65534) =>
	['ColorSchemeIndex', 'EffectSchemeIndex', 'ConnectorSchemeIndex']
		.map((name) => cell(name, scheme))
		.join('') +
	['VariationColorIndex', 'VariationStyleIndex'].map((name) => cell(name, 65534)).join('') +
	['Fill', 'Line', 'Font']
		.map((kind) => cell(`QuickStyle${kind}Color`, 100) + cell(`QuickStyle${kind}Matrix`, 100))
		.join('');
export async function themeFixture(
	options: {
		contents?: string;
		pageCells?: string;
		theme?: string;
		omitTheme?: boolean;
		external?: boolean;
	} = {},
): Promise<Uint8Array> {
	return fixture({
		document: `<StyleSheets><StyleSheet ID="0">${themeSelectors()}${cell('FillForegnd', 'Themed', 'THEMEVAL()')}${cell('LineColor', 'Themed', 'THEMEVAL()')}<Section N="Character"><Row IX="0">${cell('Color', 'Themed', 'THEMEVAL()')}</Row></Section></StyleSheet></StyleSheets>`,
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', rectangle + (options.contents ?? ''))}</Shapes>` },
		],
		edit: (zip) => {
			zip.file(
				'visio/pages/pages.xml',
				xml(
					'Pages',
					`<Page ID="0"><PageSheet>${cell('PageWidth', 8)}${cell('PageHeight', 10)}${cell('ColorSchemeIndex', 42)}${cell('EffectSchemeIndex', 42)}${cell('ConnectorSchemeIndex', 42)}${options.pageCells ?? ''}</PageSheet><Rel r:id="rId1"/></Page>`,
				),
			);
			if (!options.omitTheme) {
				zip.file(
					'visio/_rels/document.xml.rels',
					relations(
						relation('rId1', 'pages', 'pages/pages.xml') +
							`<Relationship Id="theme" Type="${themeRelationship}" Target="${options.external ? 'https://example.invalid/theme.xml' : 'theme/theme1.xml'}"${options.external ? ' TargetMode="External"' : ''}/>`,
					),
				);
				if (!options.external)
					zip.file('visio/theme/theme1.xml', options.theme ?? generatedTheme());
			}
		},
	});
}
