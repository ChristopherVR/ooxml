import { describe, expect, it } from 'vitest';
import { NS, parseXml, buildXml } from '../xml/index';
import { parseDrawingBodyProperties } from './read-text-body';
import { writeDrawingBodyProperties } from './write-text-body';
const source = `<a:bodyPr xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:scene3d v:keep="scene"><a:camera prst="perspectiveFront" zoom="125%" fov="3000000"><a:rot lat="0" lon="60000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/><a:extLst><a:ext uri="keep"><v:opaque/></a:ext></a:extLst></a:scene3d><a:sp3d extrusionH="12700" contourW="200" z="0" prstMaterial="plastic" v:keep="shape"><a:bevelT prst="circle"/><a:extrusionClr><a:schemeClr val="accent2"><a:alpha val="50000"/></a:schemeClr></a:extrusionClr><a:extLst><a:ext uri="shape"><v:opaque/></a:ext></a:extLst></a:sp3d></a:bodyPr>`;
const read = (xml = source) => parseDrawingBodyProperties(parseXml(xml).documentElement)!;
const write = (value: ReturnType<typeof read>) =>
	buildXml(writeDrawingBodyProperties(parseXml('<r/>'), value));
describe('source-backed typed text-body 3D', () => {
	it('reads scene, material, bevel and theme colour without inventing absent dimensions', () => {
		const value = read();
		expect(value.scene3d).toMatchObject({
			cameraPreset: 'perspectiveFront',
			cameraZoom: 1.25,
			cameraFieldOfView: 3000000,
			cameraRotY: 60000,
			lightRigType: 'threePt',
		});
		expect(value.shape3d).toMatchObject({
			extrusionHeight: 12700,
			contourWidth: 200,
			z: 0,
			presetMaterial: 'plastic',
			bevelTopType: 'circle',
			extrusionColor: { kind: 'scheme', value: 'accent2' },
		});
		expect(value.shape3d?.bevelTopWidth).toBeUndefined();
		expect(value.extraXml).toBeUndefined();
		expect(read(write(value))).toMatchObject({
			scene3d: { cameraZoom: 1.25 },
			shape3d: { extrusionHeight: 12700 },
		});
	});
	it('patches supported fields while retaining authored percentages, theme tokens and opaque children', () => {
		const value = read();
		value.shape3d!.extrusionHeight = 25400;
		value.shape3d!.bevelTopWidth = 0;
		value.scene3d!.cameraRotY = 120000;
		const xml = write(value);
		expect(xml).toContain('extrusionH="25400"');
		expect(xml).toContain('w="0"');
		expect(xml).toContain('zoom="125%"');
		expect(xml).toContain('lon="120000"');
		expect(xml).toContain('v:keep="shape"');
		expect(xml).toContain('val="accent2"');
		expect(xml.match(/v:opaque/g)).toHaveLength(2);
	});
	it('preserves malformed source, rejects invalid edits and maintains the flat/extruded choice', () => {
		const invalid = read(
			source
				.replace('zoom="125%"', 'zoom="broken"')
				.replace('prst="perspectiveFront"', 'prst="vendorCamera"'),
		);
		expect(write(invalid)).toContain('zoom="broken"');
		expect(write(invalid)).toContain('prst="vendorCamera"');
		const value = read();
		value.scene3d!.cameraZoom = 0;
		expect(() => write(value)).toThrow(RangeError);
		const flat = read(`<a:bodyPr xmlns:a="${NS.a}"><a:flatTx z="0"/></a:bodyPr>`);
		expect(flat.flatText3d?.z).toBe(0);
		expect(write(flat)).toContain('<a:flatTx');
		flat.shape3d = { extrusionHeight: 1 };
		expect(() => write(flat)).toThrow(/mutually exclusive/);
	});
	it('inserts new bevel children before carried colour and extension nodes', () => {
		const value = read();
		value.shape3d!.bevelBottomType = 'angle';
		const xml = write(value);
		expect(xml.indexOf('<a:bevelB')).toBeLessThan(xml.indexOf('<a:extrusionClr'));
	});
});

it.each(['cameraPreset', 'lightRigType', 'lightRigDirection'] as const)(
	'rejects removal of required scene attribute %s',
	(key) => {
		const value = read();
		delete value.scene3d![key];
		expect(() => write(value)).toThrow(/required/);
	},
);
it('keeps malformed rotation opaque through another axis edit and rejects invalid new angles', () => {
	const value = read(source.replace('lat="0"', 'lat="-1"'));
	expect(value.scene3d?.cameraRotX).toBeUndefined();
	expect(write(value)).toContain('lat="-1"');
	value.scene3d!.cameraRotY = 123;
	expect(write(value)).toContain('lat="-1"');
	for (const angle of [-1, 21600000, 0.5]) {
		const edited = read();
		edited.scene3d!.cameraRotX = angle;
		expect(() => write(edited)).toThrow(RangeError);
	}
});
it.each(['scene3d', 'sp3d'] as const)(
	'carries duplicate %s children opaquely without dropping either',
	(name) => {
		const value = read(
			`<a:bodyPr xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:${name} v:identity="first"/><a:${name} v:identity="second"/></a:bodyPr>`,
		);
		expect(name === 'scene3d' ? value.scene3d : value.shape3d).toBeUndefined();
		const xml = write(value);
		expect(xml).toContain('v:identity="first"');
		expect(xml).toContain('v:identity="second"');
	},
);
