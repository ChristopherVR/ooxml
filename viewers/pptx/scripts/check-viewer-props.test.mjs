import assert from 'node:assert/strict';
import { test } from 'node:test';

import {
	diffViewerSurface,
	VIEWER_BINDINGS,
} from '../../../src/ui/src/pptx/render/viewer-props-manifest.ts';
import { findViewerPropDrift, reportViewerPropDrift } from './check-viewer-props.mjs';
import { extractBindings } from './viewer-props-extract.mjs';

const clone = () => structuredClone(extractBindings());
const messages = (extracted) =>
	diffViewerSurface(extracted).map((d) => `${d.binding}: ${d.message}`);

test('the five bindings match the manifest today', () => {
	assert.deepEqual(findViewerPropDrift(), []);
});

test('every binding is read from source', () => {
	const extracted = extractBindings();
	for (const binding of VIEWER_BINDINGS) {
		assert.ok(extracted[binding].props.length > 10, `${binding} props`);
		assert.ok(extracted[binding].events.length > 5, `${binding} events`);
	}
});

test('a prop removed from one binding is named with that binding', () => {
	const extracted = clone();
	extracted.svelte.props = extracted.svelte.props.filter((name) => name !== 'authorName');
	const found = messages(extracted);
	assert.equal(found.length, 1);
	assert.match(found[0], /^svelte: svelte lacks the prop\/input\/option "authorName"/u);
});

test('an event removed from one binding is named with that binding', () => {
	const extracted = clone();
	extracted.vue.events = extracted.vue.events.filter((name) => name !== 'load');
	const found = messages(extracted);
	assert.equal(found.length, 1);
	assert.match(found[0], /^vue: vue lacks the event "load"/u);
});

test('a prop the manifest does not know is reported', () => {
	const extracted = clone();
	extracted.angular.props.push('brandNewInput');
	assert.match(
		messages(extracted).join('\n'),
		/angular exposes the prop\/input\/option "brandNewInput"/u,
	);
});

test('a stale exception is reported', () => {
	const extracted = clone();
	extracted.react.props.push('showNotes');
	assert.match(
		messages(extracted).join('\n'),
		/react exposes .* "showNotes" although the manifest lists it as absent/u,
	);
});

test('the report lists each difference', () => {
	const extracted = clone();
	extracted.vanilla.props = extracted.vanilla.props.filter((name) => name !== 'filePath');
	const report = reportViewerPropDrift(diffViewerSurface(extracted));
	assert.match(report, /1 difference between/u);
	assert.match(report, /\[vanilla\]/u);
});
