import {
	applyDrawingColorTransforms,
	resolveObjectColorChoice,
} from '../../color/color-transforms';
import type { XmlObject } from '../../types';

export interface PptxColorTransformCodecContext {
	resolveThemeColor: (schemeKey: string) => string | undefined;
}

export interface IPptxColorTransformCodec {
	parseColorChoice(
		colorChoice: XmlObject | undefined,
		placeholderColor?: string,
	): string | undefined;
	parseColor(colorNode: XmlObject | undefined, placeholderColor?: string): string | undefined;
	percentAttrToUnit(value: unknown): number | undefined;
	clampUnitInterval(value: number): number;
	hexToRgb(hex: string): { r: number; g: number; b: number } | undefined;
	rgbToHex(r: number, g: number, b: number): string;
	applyColorTransforms(baseColor: string, colorNode: XmlObject): string;
}

export class PptxColorTransformCodec implements IPptxColorTransformCodec {
	private readonly context: PptxColorTransformCodecContext;

	public constructor(context: PptxColorTransformCodecContext) {
		this.context = context;
	}

	public parseColorChoice(
		colorChoice: XmlObject | undefined,
		placeholderColor?: string,
	): string | undefined {
		return resolveObjectColorChoice(
			colorChoice,
			(schemeKey) => this.context.resolveThemeColor(schemeKey),
			placeholderColor,
		);
	}

	public parseColor(
		colorNode: XmlObject | undefined,
		placeholderColor?: string,
	): string | undefined {
		if (!colorNode) {
			return undefined;
		}
		return this.parseColorChoice(colorNode, placeholderColor);
	}

	public percentAttrToUnit(value: unknown): number | undefined {
		const parsedValue = Number.parseFloat(String(value ?? '').trim());
		if (!Number.isFinite(parsedValue)) {
			return undefined;
		}
		return Math.min(1, Math.max(0, parsedValue / 100000));
	}

	public clampUnitInterval(value: number): number {
		return Math.min(1, Math.max(0, value));
	}

	public hexToRgb(hex: string): { r: number; g: number; b: number } | undefined {
		const normalized = hex.replace('#', '');
		if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
			return undefined;
		}
		return {
			r: Number.parseInt(normalized.slice(0, 2), 16),
			g: Number.parseInt(normalized.slice(2, 4), 16),
			b: Number.parseInt(normalized.slice(4, 6), 16),
		};
	}

	public rgbToHex(r: number, g: number, b: number): string {
		const toHex = (channelValue: number): string =>
			this.clampColorChannel(channelValue).toString(16).padStart(2, '0').toUpperCase();
		return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
	}

	public applyColorTransforms(baseColor: string, colorNode: XmlObject): string {
		return applyDrawingColorTransforms(baseColor, colorNode);
	}

	private clampColorChannel(value: number): number {
		return Math.min(255, Math.max(0, Math.round(value)));
	}
}
