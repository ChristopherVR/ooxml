import { type EmfAdmissionContext as Context } from './emf-admission-context.js';

export const EMF_FIXED_SIZES = new Map<number, number>([
	[9, 16],
	[10, 16],
	[11, 16],
	[12, 16],
	[17, 12],
	[18, 12],
	[19, 12],
	[20, 12],
	[22, 12],
	[24, 12],
	[25, 12],
	[27, 16],
	[30, 24],
	[33, 8],
	[34, 12],
	[37, 12],
	[38, 28],
	[39, 24],
	[40, 12],
	[42, 24],
	[43, 24],
	[45, 40],
	[47, 40],
	[48, 12],
	[54, 16],
]);

// Only these text state records are inert because every text/font drawing record is rejected.
export function inspectEmfStateRecord(c: Context): boolean {
	const t = c.type;
	if ((t === 9 || t === 10 || t === 11 || t === 12 || t === 17) && c.stack.length)
		c.emit(
			'unsupported',
			'mapping-restore-fidelity',
			'Mapping mutations in saved DC state require corrected converter restoration.',
		);
	if (t === 33) {
		const depth = c.stack.length + 1;
		c.metrics.peakStateDepth = Math.max(c.metrics.peakStateDepth, depth);
		if (
			c.bound(depth, c.limits.maxStateDepth, 'state-depth', 'SaveDC depth exceeds budget.') &&
			c.charge(16)
		)
			c.stack.push({ ...c.state });
	} else if (t === 34) {
		const relative = c.i32(8);
		if (relative >= 0 || -relative > c.stack.length) {
			if (c.semanticStateKnown)
				c.emit(
					'invalid',
					'restore-state',
					'RestoreDC must reference an existing relative saved state.',
				);
		} else {
			c.state = c.stack[c.stack.length + relative]!;
			c.stack.length += relative;
		}
	} else if (t === 17) {
		if (c.state.mapMode === 8 && c.u32(8) === 1)
			c.emit(
				'unsupported',
				'mapping-transition-fidelity',
				'Returning from anisotropic mapping to MM_TEXT requires corrected converter semantics.',
			);
		if (![1, 8].includes(c.u32(8))) {
			c.emit('unsupported', 'mapping-mode', 'Only MM_TEXT and MM_ANISOTROPIC are characterized.');
			c.semanticStateKnown = false;
		} else c.state.mapMode = c.u32(8);
	} else if (t === 9 || t === 10 || t === 11 || t === 12) {
		const x = c.i32(8),
			y = c.i32(12),
			s = c.state;
		c.point(x, y, false);
		if (t === 10) {
			s.windowOriginSet = true;
			s.windowX = x;
			s.windowY = y;
		}
		if (t === 12) {
			s.viewportX = x;
			s.viewportY = y;
		}
		if (t === 9 || t === 11) {
			if (x <= 0 || y <= 0) {
				c.emit('unsupported', 'parameter', 'Nonpositive mapping extents are outside this subset.');
				return true;
			}
			if (s.mapMode === 1) {
				// MS-EMF ignores extents in MM_TEXT; the audited converter did not.
				c.emit(
					'unsupported',
					'mapping-extents-fidelity',
					'MM_TEXT extent records require corrected converter semantics.',
				);
				return true;
			}
			if (t === 9) {
				s.windowExtentSet = true;
				s.windowWidth = x;
				s.windowHeight = y;
			} else {
				s.viewportExtentSet = true;
				s.viewportWidth = x;
				s.viewportHeight = y;
			}
			for (const scale of [
				s.viewportWidth / s.windowWidth,
				s.viewportHeight / s.windowHeight,
				s.viewportWidth / s.windowWidth / (s.viewportHeight / s.windowHeight),
			]) {
				c.bound(scale, c.limits.maxMappingScale, 'mapping-scale', 'Mapping scale exceeds budget.');
				c.bound(
					1 / scale,
					c.limits.maxMappingScale,
					'mapping-scale',
					'Inverse mapping scale exceeds budget.',
				);
			}
		}
	} else if (t === 18 || t === 19) {
		if (![1, 2].includes(c.u32(8)))
			c.emit('invalid', 'parameter', 'Invalid background or polygon fill mode.');
	} else if (t === 20) {
		if (c.u32(8) !== 13) c.emit('unsupported', 'parameter', 'Only copy-pen ROP2 is admitted.');
	} else if (t === 22) {
		if (![0, 24].includes(c.u32(8)))
			c.emit('unsupported', 'parameter', 'Text alignment is outside the inert-state subset.');
	} else if (t === 24 || t === 25) inspectEmfColor(c, c.u32(8));
	else if (t === 48) {
		if (c.u32(8) !== 0x8000000f)
			c.emit('unsupported', 'palette', 'Only DEFAULT_PALETTE selection is characterized.');
	} else return false;
	return true;
}

export function inspectEmfColor(c: Context, color: number): void {
	if (color >>> 24)
		c.emit(
			'unsupported',
			'palette-color',
			'Palette-relative colors require an explicit output-device compatibility contract.',
		);
}

export function inspectEmfObjectRecord(c: Context): boolean {
	const t = c.type;
	if (t !== 37 && t !== 38 && t !== 39 && t !== 40) return false;
	const handle = c.u32(8);
	if (t === 38 || t === 39) {
		c.metrics.objectsCreated++;
		c.bound(
			c.metrics.objectsCreated,
			c.limits.maxObjects,
			'object-limit',
			'Aggregate object creation budget exceeded.',
		);
		const valid = handle > 0 && handle < (c.header?.declaredHandles ?? 0) && !c.objects.has(handle);
		if (c.semanticStateKnown && !valid)
			c.emit(
				'invalid',
				'object-handle',
				'Object handle is reserved, duplicated, or outside the declared table.',
			);
		if (valid && !c.budgetExceeded && c.charge(8))
			c.objects.set(handle, { kind: t === 38 ? 'pen' : 'brush', width: t === 38 ? c.i32(16) : 0 });
		c.metrics.peakLiveObjects = Math.max(c.metrics.peakLiveObjects, c.objects.size);
		const style = c.u32(12);
		if (t === 38) {
			if (![0, 5].includes(style))
				c.emit(
					'unsupported',
					'pen-style',
					'Only solid/null pens are admitted; INSIDEFRAME is not a centered stroke.',
				);
			const width = c.i32(16);
			if (width < 0 || c.i32(20) !== 0)
				c.emit('invalid', 'parameter', 'Pen width must be nonnegative with zero Y component.');
			c.bound(width, c.limits.maxPenWidth, 'coordinate-limit', 'Pen width exceeds budget.');
			// A zero width is a device hairline, not zero geometry. Conversion remains disabled.
			inspectEmfColor(c, c.u32(24));
		} else {
			if (![0, 1].includes(style))
				c.emit('unsupported', 'brush-style', 'Only solid/null brushes are admitted.');
			inspectEmfColor(c, c.u32(16));
			if (c.u32(20) !== 0)
				c.emit('unsupported', 'brush-style', 'Nonzero brush hatch data is outside this subset.');
		}
	} else if (t === 37) {
		const kind =
			c.objects.get(handle)?.kind ??
			(handle >= 0x80000000 && handle <= 0x80000005
				? 'brush'
				: handle >= 0x80000006 && handle <= 0x80000008
					? 'pen'
					: undefined);
		if (!kind && c.semanticStateKnown)
			c.emit('unsupported', 'object-reference', 'Unknown or unsupported object selection.');
		else if (kind) c.state[kind] = handle;
	} else if (c.semanticStateKnown) {
		c.charge(c.stack.length + 1);
		if (!c.objects.has(handle))
			c.emit('invalid', 'object-reference', 'DeleteObject references an absent object.');
		else if ([c.state, ...c.stack].some((s) => s.pen === handle || s.brush === handle))
			c.emit(
				'unsupported',
				'selected-object',
				'Deletion of an object referenced by live or saved state is unresolved.',
			);
		else c.objects.delete(handle);
	}
	return true;
}
