import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { readDeclaredFonts } from './declared-fonts';

const zip = (parts: Record<string, string>) =>
	zipSync(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)])));

const theme = (major: string, minor: string) =>
	`<a:theme><a:fontScheme><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/></a:minorFont></a:fontScheme></a:theme>`;

describe('readDeclaredFonts', () => {
	it('reads Word font table, styles and theme, theme first', () => {
		const fonts = readDeclaredFonts(
			zip({
				'word/document.xml':
					'<w:document><w:r><w:rFonts w:ascii="Ignored Body"/></w:r></w:document>',
				'word/fontTable.xml':
					'<w:fonts><w:font w:name="Lobster"><w:panose1/></w:font><w:font w:name="Corp Sans"><w:embedRegular r:id="rId1"/></w:font></w:fonts>',
				'word/styles.xml':
					'<w:styles><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:asciiTheme="minorHAnsi"/></w:styles>',
				'word/theme/theme1.xml': theme('Aptos Display', 'Aptos'),
			}),
		);
		expect(fonts.families).toEqual(['Aptos Display', 'Aptos', 'Lobster', 'Georgia']);
		expect(fonts.embedded).toEqual(['Corp Sans']);
	});

	it('reads Excel style fonts, PowerPoint parts and Visio face names', () => {
		expect(
			readDeclaredFonts(
				zip({ 'xl/styles.xml': '<fonts><font><sz val="11"/><name val="Calibri"/></font></fonts>' }),
			).families,
		).toEqual(['Calibri']);
		const deck = readDeclaredFonts(
			zip({
				'ppt/presentation.xml':
					'<p:presentation><p:embeddedFontLst><p:embeddedFont><p:font typeface="Lato"/><p:regular r:id="rId9"/></p:embeddedFont></p:embeddedFontLst></p:presentation>',
				'ppt/slides/slide1.xml':
					'<p:sld><a:latin typeface="Lato"/><a:latin typeface="Roboto"/><a:latin typeface="+mn-lt"/></p:sld>',
				'ppt/slides/slide99.xml': '<p:sld><a:latin typeface="Far Away"/></p:sld>',
				'ppt/theme/theme1.xml': theme('Montserrat', 'Open Sans'),
			}),
		);
		expect(deck.families).toEqual(['Montserrat', 'Open Sans', 'Roboto']);
		expect(deck.embedded).toEqual(['Lato']);
		expect(
			readDeclaredFonts(
				zip({
					'visio/document.xml':
						'<FaceNames><FaceName NameU="Segoe UI" Name="Segoe UI"/></FaceNames>',
				}),
			).families,
		).toEqual(['Segoe UI']);
	});

	it('gives nothing for bytes that are not a zip, and decodes entities', () => {
		expect(readDeclaredFonts(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]))).toEqual({
			families: [],
			embedded: [],
		});
		expect(readDeclaredFonts(zip({ 'xl/styles.xml': '<name val="A &amp; B"/>' })).families).toEqual(
			['A & B'],
		);
	});
});
