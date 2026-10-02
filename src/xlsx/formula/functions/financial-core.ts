// Time-value-of-money formulas shared by the financial functions.
import { ERR, fail } from '../values.js';

export function pmt(rate: number, nper: number, pv: number, fv = 0, type = 0): number {
	if (nper === 0) fail(ERR.NUM);
	if (rate === 0) return -(pv + fv) / nper;
	const f = Math.pow(1 + rate, nper);
	return (-rate * (fv + pv * f)) / ((1 + rate * type) * (f - 1));
}

export function fv(rate: number, nper: number, payment: number, pv = 0, type = 0): number {
	if (rate === 0) return -(pv + payment * nper);
	const f = Math.pow(1 + rate, nper);
	return -(pv * f + (payment * (1 + rate * type) * (f - 1)) / rate);
}

export function pv(rate: number, nper: number, payment: number, future = 0, type = 0): number {
	if (rate === 0) return -(future + payment * nper);
	const f = Math.pow(1 + rate, nper);
	return -(future + (payment * (1 + rate * type) * (f - 1)) / rate) / f;
}

export function ipmt(
	rate: number,
	per: number,
	nper: number,
	present: number,
	future = 0,
	type = 0,
): number {
	if (per < 1 || per > nper) fail(ERR.NUM);
	const payment = pmt(rate, nper, present, future, type);
	let interest: number;
	if (per === 1) interest = type === 1 ? 0 : -present;
	else
		interest =
			type === 1
				? fv(rate, per - 2, payment, present, 1) - payment
				: fv(rate, per - 1, payment, present, 0);
	return interest * rate;
}

/** Newton's method with a bisection fallback; `#NUM!` when no root is found. */
export function solve(f: (x: number) => number, guess: number): number {
	let x = guess;
	for (let i = 0; i < 100; i++) {
		const y = f(x);
		if (!Number.isFinite(y)) break;
		if (y === 0) return x;
		const h = Math.max(1e-9, Math.abs(x) * 1e-9);
		const slope = (f(x + h) - f(x - h)) / (2 * h);
		if (slope === 0 || !Number.isFinite(slope)) break;
		const next = x - y / slope;
		if (Math.abs(next - x) <= 1e-14 * Math.max(1, Math.abs(next))) return next;
		x = next;
		if (x <= -1) x = -0.999999;
	}
	let lo = -0.999999;
	let hi = 10;
	let flo = f(lo);
	if (!Number.isFinite(flo) || flo * f(hi) > 0) fail(ERR.NUM);
	for (let i = 0; i < 300; i++) {
		const mid = (lo + hi) / 2;
		const fm = f(mid);
		if (Math.abs(fm) < 1e-10) return mid;
		if (fm * flo < 0) hi = mid;
		else {
			lo = mid;
			flo = fm;
		}
	}
	return (lo + hi) / 2;
}

export function cumulative(args: number[], principal: boolean): number {
	const [rate = 0, nper = 0, present = 0, start = 0, end = 0, type = 0] = args;
	if (
		rate <= 0 ||
		nper <= 0 ||
		present <= 0 ||
		start < 1 ||
		end < start ||
		end > nper ||
		(type !== 0 && type !== 1)
	)
		fail(ERR.NUM);
	let total = 0;
	const payment = pmt(rate, nper, present, 0, type);
	for (let per = Math.trunc(start); per <= Math.trunc(end); per++) {
		const interest = ipmt(rate, per, nper, present, 0, type);
		total += principal ? payment - interest : interest;
	}
	return total;
}
