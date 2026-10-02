import {
	VISIO_EMF_ADMISSION_LIMITS,
	type VisioEmfAdmissionOptions,
	type VisioEmfAdmissionResult,
	type VisioEmfAdmissionDiagnostic,
	type VisioEmfAdmissionHeader,
	type VisioEmfAdmissionMetrics,
	type VisioEmfDiagnosticCode,
} from './emf-admission-types.js';

export interface EmfState {
	mapMode: number;
	windowExtentSet: boolean;
	viewportExtentSet: boolean;
	windowOriginSet: boolean;
	windowX: number;
	windowY: number;
	viewportX: number;
	viewportY: number;
	windowWidth: number;
	windowHeight: number;
	viewportWidth: number;
	viewportHeight: number;
	pen: number;
	brush: number;
	x: number;
	y: number;
	clipDepth: number;
}
export class EmfAdmissionContext {
	readonly limits: { -readonly [K in keyof typeof VISIO_EMF_ADMISSION_LIMITS]: number };
	readonly diagnostics: VisioEmfAdmissionDiagnostic[] = [];
	readonly inventory = new Map<number, number>();
	readonly objects = new Map<number, { kind: 'pen' | 'brush'; width: number }>();
	readonly stack: EmfState[] = [];
	readonly metrics: VisioEmfAdmissionMetrics = {
		records: 0,
		objectsCreated: 0,
		peakLiveObjects: 0,
		points: 0,
		regionRects: 0,
		clipOperations: 0,
		peakClipDepth: 0,
		peakStateDepth: 0,
		paletteEntries: 0,
		commentBytes: 0,
		allocationUnits: 0,
		maxMappedCoordinate: 0,
	};
	state: EmfState = {
		mapMode: 1,
		windowExtentSet: false,
		viewportExtentSet: false,
		windowOriginSet: false,
		windowX: 0,
		windowY: 0,
		viewportX: 0,
		viewportY: 0,
		windowWidth: 1,
		windowHeight: 1,
		viewportWidth: 1,
		viewportHeight: 1,
		pen: 0x80000007,
		brush: 0x80000000,
		x: 0,
		y: 0,
		clipDepth: 0,
	};
	header: VisioEmfAdmissionHeader | null = null;
	headerPaletteEntries = 0;
	offset = 0;
	type: number | null = null;
	scanComplete = false;
	semanticStateKnown = true;
	omittedDiagnostics = 0;
	budgetExceeded = false;
	invalid = false;
	unsupported = false;
	constructor(
		readonly view: DataView,
		options: VisioEmfAdmissionOptions,
	) {
		this.limits = { ...VISIO_EMF_ADMISSION_LIMITS };
		if (!options || typeof options !== 'object') {
			this.emit('invalid', 'limits', 'Expected an options object.');
			return;
		}
		const cap = options.maxDiagnostics;
		if (cap !== undefined && Number.isInteger(cap) && cap > 0 && cap <= this.limits.maxDiagnostics)
			this.limits.maxDiagnostics = cap;
		for (const key of Object.keys(VISIO_EMF_ADMISSION_LIMITS) as (keyof typeof this.limits)[]) {
			const value = options[key];
			if (value === undefined) continue;
			if (!Number.isSafeInteger(value) || value < 1 || value > this.limits[key])
				this.emit('invalid', 'limits', 'Limits must be positive integers below the hard ceiling.');
			else this.limits[key] = value;
		}
	}
	u32(at: number): number {
		return this.view.getUint32(this.offset + at, true);
	}
	i32(at: number): number {
		return this.view.getInt32(this.offset + at, true);
	}
	i16(at: number): number {
		return this.view.getInt16(this.offset + at, true);
	}
	emit(
		kind: VisioEmfAdmissionDiagnostic['kind'],
		code: VisioEmfDiagnosticCode,
		message: string,
		limit?: number,
	): void {
		if (kind === 'invalid') this.invalid = true;
		if (kind === 'unsupported') this.unsupported = true;
		if (kind === 'budget-exceeded') this.budgetExceeded = true;
		if (this.diagnostics.length < this.limits.maxDiagnostics)
			this.diagnostics.push({
				kind,
				code,
				offset: this.offset,
				recordType: this.type,
				message,
				...(limit === undefined ? {} : { limit }),
			});
		else this.omittedDiagnostics++;
	}
	charge(units: number): boolean {
		this.metrics.allocationUnits += units;
		return this.bound(
			this.metrics.allocationUnits,
			this.limits.maxAllocationUnits,
			'allocation-limit',
			'Aggregate abstract allocation/work budget exceeded.',
		);
	}
	bound(value: number, limit: number, code: VisioEmfDiagnosticCode, message: string): boolean {
		if (Number.isFinite(value) && Math.abs(value) <= limit) return true;
		this.emit('budget-exceeded', code, message, limit);
		return false;
	}
	coordinate(value: number): boolean {
		return this.bound(
			value,
			this.limits.maxCoordinate,
			'coordinate-limit',
			'Source coordinate exceeds budget.',
		);
	}
	point(x: number, y: number, mapped = true): void {
		this.coordinate(x);
		this.coordinate(y);
		if (!mapped || !this.semanticStateKnown) return;
		const s = this.state;
		if (
			s.mapMode === 8 &&
			(!s.windowExtentSet ||
				!s.viewportExtentSet ||
				(!s.windowOriginSet && (this.header?.bounds[0] !== 0 || this.header?.bounds[1] !== 0)))
		)
			this.emit(
				'unsupported',
				'mapping-incomplete',
				'Anisotropic geometry requires explicit extents and a resolved window origin.',
			);
		const sx = s.mapMode === 8 ? s.viewportWidth / s.windowWidth : 1;
		const sy = s.mapMode === 8 ? s.viewportHeight / s.windowHeight : 1;
		const mx = (x - s.windowX) * sx + s.viewportX;
		const my = (y - s.windowY) * sy + s.viewportY;
		this.metrics.maxMappedCoordinate = Math.max(
			this.metrics.maxMappedCoordinate,
			Math.abs(mx),
			Math.abs(my),
		);
		this.bound(
			mx,
			this.limits.maxCoordinate,
			'coordinate-limit',
			'Mapped coordinate exceeds budget.',
		);
		this.bound(
			my,
			this.limits.maxCoordinate,
			'coordinate-limit',
			'Mapped coordinate exceeds budget.',
		);
	}
	rect(at: number, mapped = true): void {
		this.point(this.i32(at), this.i32(at + 4), mapped);
		this.point(this.i32(at + 8), this.i32(at + 12), mapped);
		if (this.i32(at) > this.i32(at + 8) || this.i32(at + 4) > this.i32(at + 12))
			this.emit(
				'unsupported',
				'parameter',
				'Inverted rectangle semantics are outside this subset.',
			);
	}
	clip(reset = false): void {
		this.metrics.clipOperations++;
		this.state.clipDepth = reset ? 0 : this.state.clipDepth + 1;
		this.metrics.peakClipDepth = Math.max(this.metrics.peakClipDepth, this.state.clipDepth);
		this.bound(
			this.metrics.clipOperations,
			this.limits.maxClipOperations,
			'clip-limit',
			'Aggregate clipping budget exceeded.',
		);
		this.bound(
			this.state.clipDepth,
			this.limits.maxClipDepth,
			'clip-limit',
			'Active clip depth exceeds budget.',
		);
		this.charge(8);
	}
	draw(points: number): void {
		const width = this.objects.get(this.state.pen)?.width ?? 1;
		if (width > 0 && this.state.mapMode === 8) {
			for (const scale of [
				this.state.viewportWidth / this.state.windowWidth,
				this.state.viewportHeight / this.state.windowHeight,
			])
				this.bound(
					width * scale,
					this.limits.maxCoordinate,
					'coordinate-limit',
					'Mapped pen width exceeds budget.',
				);
		}
		this.metrics.points += points;
		this.bound(
			this.metrics.points,
			this.limits.maxPoints,
			'point-limit',
			'Aggregate point budget exceeded.',
		);
		// Charge repeated traversal of the complete active clip ancestry at every draw.
		this.charge((points + 16) * (this.state.clipDepth + 1));
	}
	result(): VisioEmfAdmissionResult {
		return {
			status: this.budgetExceeded
				? 'budget-exceeded'
				: this.invalid
					? 'invalid'
					: this.unsupported
						? 'unsupported'
						: 'admitted',
			renderingEnabled: false,
			header: this.header,
			metrics: this.metrics,
			recordTypes: Array.from(this.inventory, ([type, count]) => ({ type, count })),
			diagnostics: this.diagnostics,
			omittedDiagnostics: this.omittedDiagnostics,
			scanComplete: this.scanComplete,
		};
	}
}
