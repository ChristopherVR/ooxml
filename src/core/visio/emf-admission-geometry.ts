import { type EmfAdmissionContext as Context } from './emf-admission-context.js';

export function inspectEmfGeometry(c: Context, size: number): boolean {
	const t = c.type;
	if (t === 27 || t === 54) {
		const x = c.i32(8),
			y = c.i32(12);
		c.point(x, y);
		if (t === 54) {
			c.point(c.state.x, c.state.y);
			c.draw(2);
		}
		c.state.x = x;
		c.state.y = y;
	} else if (t === 30) {
		c.rect(8);
		c.clip();
	} else if (t === 42 || t === 43 || t === 45 || t === 47) {
		c.rect(8);
		if (c.i32(8) === c.i32(16) || c.i32(12) === c.i32(20))
			c.emit('unsupported', 'parameter', 'Degenerate bounded primitives are outside this subset.');
		if (t === 45 || t === 47) {
			// Radial direction points may be outside the ellipse. They are not arc endpoints.
			c.point(c.i32(24), c.i32(28));
			c.point(c.i32(32), c.i32(36));
			const cx = (c.i32(8) + c.i32(16)) / 2,
				cy = (c.i32(12) + c.i32(20)) / 2;
			if ((c.i32(24) === cx && c.i32(28) === cy) || (c.i32(32) === cx && c.i32(36) === cy))
				c.emit(
					'unsupported',
					'parameter',
					'An arc radial point at the center has unresolved semantics.',
				);
			// Twice the center keeps odd-size ellipses integral. Coordinate ceilings make
			// these products exact (< 2^53). Same ray is a full ellipse, not an empty arc.
			if (!c.budgetExceeded) {
				const ax = 2 * c.i32(24) - c.i32(8) - c.i32(16),
					ay = 2 * c.i32(28) - c.i32(12) - c.i32(20);
				const bx = 2 * c.i32(32) - c.i32(8) - c.i32(16),
					by = 2 * c.i32(36) - c.i32(12) - c.i32(20);
				if (ax * by === ay * bx && ax * bx + ay * by > 0)
					c.emit(
						'unsupported',
						'arc-full-ellipse-fidelity',
						'Same-ray arc endpoints require full-ellipse converter semantics.',
					);
			}
		}
		c.draw(t === 43 ? 4 : 16);
	} else if (t === 86 || t === 87) {
		if (size < 28) {
			c.emit('invalid', 'payload-size', 'Truncated 16-bit poly record.');
			return true;
		}
		const count = c.u32(24);
		if (size !== 28 + count * 4 || count < (t === 86 ? 3 : 2)) {
			c.emit('invalid', 'payload-size', 'Point count must exactly match the record payload.');
			return true;
		}
		c.draw(count);
		if (c.budgetExceeded) return true;
		// Record Bounds are advisory; validate independently, never use them instead of points.
		c.rect(8, false);
		for (let i = 0; i < count; i++) c.point(c.i16(28 + i * 4), c.i16(30 + i * 4));
	} else if (t === 75) inspectRegion(c, size);
	else return false;
	return true;
}

function inspectRegion(c: Context, size: number): void {
	if (size < 16) {
		c.emit('invalid', 'payload-size', 'Truncated clip region record.');
		return;
	}
	const bytes = c.u32(8),
		mode = c.u32(12);
	if (bytes !== size - 16) {
		c.emit('invalid', 'region-data', 'Region byte count must exactly match the record.');
		return;
	}
	if (mode !== 5)
		c.emit('unsupported', 'region-mode', 'Only copy/reset clip regions are characterized.');
	if (bytes === 0) {
		if (mode !== 5) c.emit('invalid', 'region-data', 'An omitted region requires RGN_COPY.');
		c.clip(true);
		return;
	}
	if (bytes < 32) {
		c.emit('invalid', 'region-data', 'Truncated RegionDataHeader.');
		return;
	}
	const count = c.u32(24),
		payload = c.u32(28);
	if (c.u32(16) !== 32 || c.u32(20) !== 1 || payload !== count * 16 || bytes !== 32 + payload) {
		c.emit('invalid', 'region-data', 'Region header/count/rectangle payload is inconsistent.');
		return;
	}
	c.metrics.regionRects += count;
	c.bound(
		c.metrics.regionRects,
		c.limits.maxRegionRects,
		'region-limit',
		'Aggregate region rectangle budget exceeded.',
	);
	if (!c.charge(count * 4) || c.budgetExceeded) return;
	c.rect(32);
	for (let i = 0; i < count; i++) c.rect(48 + i * 16);
	// Multiple-rectangle band ordering, unions, and device/logical clipping need independent review.
	c.emit('unsupported', 'region-mode', 'Nonempty region replay semantics are not admitted.');
	c.clip();
}
