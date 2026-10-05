/** Hard ceilings for the DOM-free, vector-only admission checker, not converter memory quotas. */
export const VISIO_EMF_ADMISSION_LIMITS = Object.freeze({
	maxInputBytes: 4 * 1024 * 1024,
	maxRecords: 20_000,
	maxRecordTypes: 128,
	maxHandles: 4096,
	maxObjects: 4096,
	maxPoints: 100_000,
	maxRegionRects: 4096,
	maxClipOperations: 4096,
	maxClipDepth: 64,
	maxStateDepth: 64,
	maxPaletteEntries: 256,
	maxCommentBytes: 256 * 1024,
	maxCoordinate: 1_000_000,
	maxMappingScale: 4096,
	maxPenWidth: 1024,
	maxAllocationUnits: 500_000,
	maxDiagnostics: 64,
});
export type VisioEmfAdmissionLimits = typeof VISIO_EMF_ADMISSION_LIMITS;
export type VisioEmfAdmissionOptions = Partial<{ [K in keyof VisioEmfAdmissionLimits]: number }>;
export type VisioEmfAdmissionStatus = 'admitted' | 'unsupported' | 'invalid' | 'budget-exceeded';
export type VisioEmfDiagnosticCode =
	| 'input'
	| 'limits'
	| 'header'
	| 'record-size'
	| 'record-count'
	| 'record-limit'
	| 'unsupported-record'
	| 'payload-size'
	| 'parameter'
	| 'coordinate-limit'
	| 'mapping-scale'
	| 'mapping-mode'
	| 'mapping-extents-fidelity'
	| 'mapping-transition-fidelity'
	| 'mapping-restore-fidelity'
	| 'mapping-incomplete'
	| 'arc-full-ellipse-fidelity'
	| 'state-depth'
	| 'restore-state'
	| 'unbalanced-state'
	| 'object-limit'
	| 'object-handle'
	| 'object-reference'
	| 'selected-object'
	| 'pen-style'
	| 'brush-style'
	| 'palette-color'
	| 'palette-limit'
	| 'palette'
	| 'point-limit'
	| 'region-limit'
	| 'region-data'
	| 'region-mode'
	| 'clip-limit'
	| 'allocation-limit'
	| 'comment-range'
	| 'comment-limit'
	| 'embedded-wmf'
	| 'emf-plus'
	| 'comment-subtype'
	| 'eof'
	| 'missing-eof'
	| 'trailing-data';
export interface VisioEmfAdmissionDiagnostic {
	readonly kind: Exclude<VisioEmfAdmissionStatus, 'admitted'>;
	readonly code: VisioEmfDiagnosticCode;
	readonly offset: number;
	readonly recordType: number | null;
	readonly message: string;
	readonly limit?: number;
}
export interface VisioEmfAdmissionMetrics {
	records: number;
	objectsCreated: number;
	peakLiveObjects: number;
	/** Conservative point-equivalent units, including fixed primitive expansion allowances. */
	points: number;
	regionRects: number;
	clipOperations: number;
	peakClipDepth: number;
	peakStateDepth: number;
	paletteEntries: number;
	commentBytes: number;
	/** Bounded abstract work ledger; not an estimate or guarantee of JS/native heap bytes. */
	allocationUnits: number;
	maxMappedCoordinate: number;
}
export interface VisioEmfAdmissionHeader {
	readonly bounds: readonly [number, number, number, number];
	/** EMF header bounds are inclusive-inclusive. */
	readonly pixelWidth: number;
	readonly pixelHeight: number;
	readonly declaredRecords: number;
	readonly declaredHandles: number;
}
interface VisioEmfAdmissionBase {
	readonly header: VisioEmfAdmissionHeader | null;
	readonly metrics: Readonly<VisioEmfAdmissionMetrics>;
	readonly recordTypes: ReadonlyArray<{ readonly type: number; readonly count: number }>;
	readonly diagnostics: readonly VisioEmfAdmissionDiagnostic[];
	readonly omittedDiagnostics: number;
	/** True only when bounded framing reached the EOF; not a full conformance assertion. */
	readonly scanComplete: boolean;
}
/** Inspection is not a reusable authorization token: mutable bytes must be checked again at use. */
export type VisioEmfAdmissionResult = VisioEmfAdmissionBase &
	(
		| { readonly status: 'admitted'; readonly renderingEnabled: false }
		| {
				readonly status: 'unsupported' | 'invalid' | 'budget-exceeded';
				readonly renderingEnabled: false;
		  }
	);
