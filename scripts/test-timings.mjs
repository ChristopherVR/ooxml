import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

export function vitestTimings(report, suite) {
	if (report.success !== true || !Array.isArray(report.testResults) || !report.testResults.length) {
		throw new Error('Expected a successful Vitest JSON report with test results');
	}
	const marker = `/src/${suite}/`;
	const timings = new Map();
	for (const { name, startTime, endTime } of report.testResults) {
		const path = name.replaceAll('\\', '/');
		const markerIndex = path.lastIndexOf(marker);
		if (
			markerIndex < 0 ||
			!Number.isFinite(startTime) ||
			!Number.isFinite(endTime) ||
			endTime < startTime
		) {
			throw new Error(`Invalid ${suite} timing for ${name}`);
		}
		const file = path.slice(markerIndex + marker.length);
		timings.set(file, (timings.get(file) ?? 0) + endTime - startTime);
	}
	return Object.fromEntries(
		[...timings.keys()]
			.sort()
			.map((file) => [file, Math.max(1, Math.round(timings.get(file) / 10)) / 100]),
	);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	const [command, report, suite] = process.argv.slice(2);
	if (command !== 'update' || !report || (suite !== 'core' && suite !== 'ui')) {
		throw new Error('Usage: node scripts/test-timings.mjs update <report.json> <core|ui>');
	}
	const timings = vitestTimings(JSON.parse(readFileSync(report, 'utf8')), suite);
	writeFileSync(
		new URL(`../src/${suite}/vitest.timings.json`, import.meta.url),
		`${JSON.stringify(timings, null, '\t')}\n`,
	);
}
