// Host-supplied synchronous provider interfaces for Excel's external and integration functions.
// Core never makes outbound network requests itself: every provider is optional and unconfigured
// by default, so unconfigured functions preserve cached values on load and return `#NAME?` when
// newly entered.
import type { CellError } from '../model';
import type { Scalar } from './values';

/** Excel KPI property codes for `CUBEKPIMEMBER` (1=Value .. 6=Current Time Member). */
export type CubeKpiProperty = 1 | 2 | 3 | 4 | 5 | 6;

/** Excel sort order codes for `CUBESET` (0=none .. 6=natural descending). */
export type CubeSortOrder = 0 | 1 | 2 | 3 | 4 | 5 | 6;

/** Opaque handle representing an evaluated OLAP member (`CUBEMEMBER`, `CUBEKPIMEMBER`, `CUBERANKEDMEMBER`). */
export interface CubeMemberHandle {
	readonly kind: 'member';
	readonly connection: string;
	readonly token: string;
	readonly caption: string;
}

/** Opaque handle representing an evaluated OLAP set (`CUBESET`). */
export interface CubeSetHandle {
	readonly kind: 'set';
	readonly connection: string;
	readonly token: string;
	readonly caption: string;
}

export type CubeHandle = CubeMemberHandle | CubeSetHandle;

/** Result returned by a `CubeProvider` for member-returning functions. */
export interface CubeMemberResult {
	/** Opaque identifier passed back into `CUBEVALUE` and `CUBEMEMBERPROPERTY`. */
	readonly token: string;
	/** Display text stored in the cell. */
	readonly caption: string;
}

/** Result returned by a `CubeProvider` for `CUBESET`. */
export interface CubeSetResult {
	/** Opaque identifier passed back into `CUBESETCOUNT`, `CUBERANKEDMEMBER` and `CUBEVALUE`. */
	readonly token: string;
	/** Display text stored in the cell. */
	readonly caption: string;
}

export interface CubeSetOptions {
	readonly caption?: string;
	readonly sortOrder?: CubeSortOrder;
	readonly sortBy?: string;
}

/** An argument passed to `CubeProvider.value`: either a raw MDX expression or a resolved handle. */
export type CubeValueArg = string | CubeMemberHandle | CubeSetHandle;

/** Host provider backing all 7 `CUBE*` functions. */
export interface CubeProvider {
	member(
		connection: string,
		memberExpression: string | readonly string[],
		caption?: string,
	): CubeMemberResult | CellError;
	kpiMember(
		connection: string,
		kpiName: string,
		kpiProperty: CubeKpiProperty,
		caption?: string,
	): CubeMemberResult | CellError;
	memberProperty(connection: string, member: string | CubeMemberHandle, property: string): Scalar;
	set(
		connection: string,
		setExpression: string,
		options?: CubeSetOptions,
	): CubeSetResult | CellError;
	setCount(set: CubeSetHandle): number | CellError;
	rankedMember(
		connection: string,
		set: string | CubeSetHandle,
		rank: number,
		caption?: string,
	): CubeMemberResult | CellError;
	value(connection: string, expressions: readonly CubeValueArg[]): Scalar;
}

/** One prompt/context pair passed to `COPILOT`. */
export interface CopilotPromptPart {
	readonly prompt: string;
	readonly context?: Scalar | readonly (readonly Scalar[])[];
}

export interface CopilotRequest {
	readonly parts: readonly CopilotPromptPart[];
}

/** Host provider backing `COPILOT`. May return a scalar or a 2-D spill matrix. */
export interface CopilotProvider {
	evaluate(request: CopilotRequest): Scalar | readonly (readonly Scalar[])[];
}

/** `STOCKHISTORY` interval code: 0 = daily, 1 = weekly, 2 = monthly. */
export type StockHistoryInterval = 0 | 1 | 2;

/** `STOCKHISTORY` property code: 0=Date, 1=Close, 2=Open, 3=High, 4=Low, 5=Volume. */
export type StockHistoryProperty = 0 | 1 | 2 | 3 | 4 | 5;

export interface StockHistoryRequest {
	readonly stock: string;
	/** Inclusive start date as an Excel date serial (integer). */
	readonly startDate: number;
	/** Inclusive end date as an Excel date serial (integer). */
	readonly endDate: number;
	readonly interval: StockHistoryInterval;
	/** Whether the workbook uses the 1904 date system. */
	readonly date1904: boolean;
}

/** One historical bar returned by `StockHistoryProvider`. */
export interface StockHistoryBar {
	/** Excel date serial for this bar. */
	readonly date: number;
	readonly close?: number | null;
	readonly open?: number | null;
	readonly high?: number | null;
	readonly low?: number | null;
	readonly volume?: number | null;
}

/** Host provider backing `STOCKHISTORY`. */
export interface StockHistoryProvider {
	history(request: StockHistoryRequest): readonly StockHistoryBar[] | CellError;
}

/** Host provider backing `TRANSLATE`. */
export interface TranslateProvider {
	translate(text: string, sourceLanguage?: string, targetLanguage?: string): string | CellError;
}

/** Host provider backing `DETECTLANGUAGE`. */
export interface DetectLanguageProvider {
	detectLanguage(text: string): string | CellError;
}

/** `IMAGE` sizing mode: 0 = fit cell, 1 = fill cell, 2 = original size, 3 = custom size. */
export type ImageSizing = 0 | 1 | 2 | 3;

export interface ImageRequest {
	readonly source: string;
	readonly altText: string;
	readonly sizing: ImageSizing;
	readonly height?: number;
	readonly width?: number;
}

/** Host provider backing `IMAGE`. */
export interface ImageProvider {
	resolve(request: ImageRequest): Scalar;
	/** Image bytes for a successfully resolved source. The engine validates them before display. */
	getImage?(source: string): { readonly bytes: Uint8Array } | undefined;
}

/** Host provider backing `RTD` (volatile when configured). */
export interface RtdProvider {
	read(progId: string, server: string, topics: readonly string[]): Scalar;
}

/** Host provider backing `WEBSERVICE`. */
export interface WebserviceProvider {
	fetch(url: string): string | CellError;
}

/** Host-supplied external function providers configured on `CalcEngineOptions`. */
export interface CalcExternalProviders {
	cube?: CubeProvider;
	copilot?: CopilotProvider;
	stockHistory?: StockHistoryProvider;
	translate?: TranslateProvider;
	detectLanguage?: DetectLanguageProvider;
	image?: ImageProvider;
	rtd?: RtdProvider;
	webservice?: WebserviceProvider;
}
