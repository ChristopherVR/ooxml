// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import {
	CONTEXTUAL_TAB_GROUPS,
	RIBBON_HOME_FAMILIES,
	arrangeShapeHomeControls,
	dispatchCropRequest,
	homeFamilyKeys,
	pictureSizeHomeControls,
	resolveCustomization,
	ribbonCustomizationCss,
} from '../render';
import type { RibbonHomeControlSpec } from '../render';
import { translationsEn } from '../i18n/translations-en';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const crop = { editable: true, canCrop: true, cropActive: true };
const specOf = (family: 'arrange-shape' | 'picture-size', id: string): RibbonHomeControlSpec =>
	RIBBON_HOME_FAMILIES[family].clusters
		.flatMap((cluster) => cluster.controls)
		.find((control) => control.id === id)!;

describe('picture Format > Size', () => {
	it('is the third Picture Format group and draws the shared picture-size strip', () => {
		const groups = CONTEXTUAL_TAB_GROUPS.pictureFormat;
		expect(groups.map((group) => group.group)).toStrictEqual([
			'pictureFormat.adjust',
			'pictureFormat.pictureStyles',
			'pictureFormat.size',
		]);
		expect(groups[2]).toMatchObject({ labelKey: 'pptx.ribbon.groupSize', strip: 'picture-size' });
		expect(translationsEn['pptx.ribbon.groupSize']).toBe('Size');
		const missing = homeFamilyKeys('picture-size').filter((key) => !(key in translationsEn));
		expect(missing).toStrictEqual([]);
	});

	it('reuses the Home Crop menu, caret and crop hooks', () => {
		const home = specOf('arrange-shape', 'home.arrange.crop');
		const size = specOf('picture-size', 'pictureFormat.size.crop');
		expect(size.items).toBe(home.items);
		expect(size.kind).toBe('menu');
		expect(size.caret?.labelKey).toBe(home.caret?.labelKey);
		expect(size.attrs?.['data-pptx-ribbon-control']).toBe('crop');
		expect(size.caret?.attrs?.['data-pptx-ribbon-control']).toBe('crop-menu');
		expect(size).toMatchObject({ large: true, icon: 'home.arrange.crop' });
	});

	it('maps one crop state onto both crop buttons', () => {
		const home = arrangeShapeHomeControls({
			...crop,
			canGroup: false,
			canUngroup: false,
			canMerge: false,
			canStrokeWidth: false,
			strokeWidth: 1,
		});
		const size = pictureSizeHomeControls(crop);
		expect(size['pictureFormat.size.crop']).toStrictEqual(home['home.arrange.crop']);
		expect(size['pictureFormat.size.crop#caret']).toStrictEqual(home['home.arrange.crop#caret']);
		expect(size['pictureFormat.size.crop']?.pressed).toBeTruthy();
		const off = pictureSizeHomeControls({ editable: true, canCrop: false, cropActive: false });
		expect(off['pictureFormat.size.crop']).toMatchObject({ disabled: true, pressed: false });
		expect(off['pictureFormat.size.crop#caret']?.disabled).toBeTruthy();
		// An open crop session keeps the button usable so it can commit the crop.
		const open = pictureSizeHomeControls({ editable: true, canCrop: false, cropActive: true });
		expect(open['pictureFormat.size.crop']?.disabled).toBeFalsy();
	});

	it('routes the toggle and each menu pick to one crop command', () => {
		const handlers = { toggle: vi.fn(), aspect: vi.fn(), fill: vi.fn(), fit: vi.fn() };
		for (const value of [undefined, 'aspect:16:9', 'fill', 'fit', 'bogus']) {
			dispatchCropRequest(value, handlers);
		}
		expect(handlers.toggle).toHaveBeenCalledOnce();
		expect(handlers.aspect).toHaveBeenCalledWith(16, 9);
		expect(handlers.fill).toHaveBeenCalledOnce();
		expect(handlers.fit).toHaveBeenCalledOnce();
	});

	it('emits the crop toggle and a picked aspect ratio under its own id', () => {
		const host = document.createElement('pptx-ui-ribbon-home-picture-size');
		host.state = { controls: pictureSizeHomeControls({ ...crop, cropActive: false }) };
		document.body.append(host);
		const request = vi.fn();
		host.addEventListener('home-request', request);
		const slot = host.querySelector<HTMLElement>(
			'[data-ribbon-control="pictureFormat.size.crop"]',
		)!;
		const main = slot.querySelector<HTMLElement>('[data-pptx-ribbon-control="crop"]')!;
		expect(main.dataset.size).toBe('large');
		expect(main.getAttribute('aria-pressed')).toBe('false');
		main.click();
		slot.querySelector<HTMLElement>('[data-pptx-ribbon-control="crop-menu"]')!.click();
		slot.querySelector<HTMLElement>('[data-pptx-crop-action="fit"]')!.click();
		expect(request.mock.calls.map((call) => call[0].detail)).toStrictEqual([
			{ id: 'pictureFormat.size.crop' },
			{ id: 'pictureFormat.size.crop', value: 'fit' },
		]);
	});
});

describe('picture Format > Size customization', () => {
	it('hides both crop entry points for the legacy crop button', () => {
		const resolved = resolveCustomization({ ribbon: { hiddenButtons: ['crop'] } });
		expect(resolved.hiddenRibbonControls.has('home.arrange.crop')).toBeTruthy();
		expect(resolved.hiddenRibbonControls.has('pictureFormat.size.crop')).toBeTruthy();
		// `home.arrange.crop` is the same control as `crop`, so the bindings hide both too.
		const home = resolveCustomization({ ribbon: { hiddenButtons: ['home.arrange.crop'] } });
		expect(home.hiddenActions.has('crop')).toBeTruthy();
	});

	it('hides only the Size group or its Crop when the host names them', () => {
		const group = resolveCustomization({ ribbon: { hiddenGroups: ['pictureFormat.size'] } });
		expect(ribbonCustomizationCss(group)).toContain('[data-ribbon-group="pictureFormat.size"]');
		expect(group.hiddenActions.has('crop')).toBeFalsy();
		const control = resolveCustomization({
			ribbon: { hiddenButtons: ['pictureFormat.size.crop'] },
		});
		expect(control.hiddenActions.has('crop')).toBeFalsy();
		expect(control.hiddenRibbonControls.has('home.arrange.crop')).toBeFalsy();
		expect(ribbonCustomizationCss(control)).toContain(
			'[data-ribbon-control="pictureFormat.size.crop"]',
		);
	});
});
