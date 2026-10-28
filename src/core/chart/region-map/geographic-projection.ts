import type { GeoPoint } from './cache';
const RAD = Math.PI / 180;
const RX = [
	1, 0.9986, 0.9954, 0.99, 0.9822, 0.973, 0.96, 0.9427, 0.9216, 0.8962, 0.8679, 0.835, 0.7986,
	0.7597, 0.7186, 0.6732, 0.6213, 0.5722, 0.5322,
];
const RY = [
	0, 0.062, 0.124, 0.186, 0.248, 0.31, 0.372, 0.434, 0.4958, 0.5571, 0.6176, 0.6769, 0.7346, 0.7903,
	0.8435, 0.8936, 0.9394, 0.9761, 1,
];
/** Spherical projections. Robinson uses five-degree table interpolation; no external basemap service. */
export function projectGeoPoint(
	point: GeoPoint,
	projection = 'mercator',
): [number, number] | undefined {
	const lon = point.longitude * RAD,
		lat = Math.max(-89.999, Math.min(89.999, point.latitude)) * RAD;
	if (projection === 'mercator')
		return [
			lon,
			-Math.log(
				Math.tan(Math.PI / 4 + Math.max(-85.05112878 * RAD, Math.min(85.05112878 * RAD, lat)) / 2),
			),
		];
	if (projection === 'miller') return [lon, -1.25 * Math.log(Math.tan(Math.PI / 4 + 0.4 * lat))];
	if (projection === 'robinson') {
		const index = Math.min(17, Math.floor(Math.abs(point.latitude) / 5)),
			t = Math.min(1, Math.abs(point.latitude) / 5 - index);
		return [
			0.8487 * lon * (RX[index]! + (RX[index + 1]! - RX[index]!) * t),
			-1.3523 * Math.sign(lat) * (RY[index]! + (RY[index + 1]! - RY[index]!) * t),
		];
	}
	if (projection === 'albers') {
		const n = (Math.sin(20 * RAD) + Math.sin(50 * RAD)) / 2;
		const c = Math.cos(20 * RAD) ** 2 + 2 * n * Math.sin(20 * RAD);
		const rho = Math.sqrt(Math.max(0, c - 2 * n * Math.sin(lat))) / n;
		return [rho * Math.sin(n * lon), rho * Math.cos(n * lon) - Math.sqrt(c) / n];
	}
	return undefined;
}
