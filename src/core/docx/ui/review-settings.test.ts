import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import {
	setReviewRecordingPreferences,
	toggleTrackChanges,
	toggleTrackFormatting,
	toggleTrackMoves,
} from './review-settings';

describe('document-wide review recording', () => {
	it('applies both preferences atomically and supports probes and no-op updates', () => {
		const schema = new Schema({
			nodes: {
				doc: {
					content: 'text*',
					attrs: { trackFormatting: { default: true }, trackMoves: { default: true } },
				},
				text: {},
			},
		});
		let state = EditorState.create({ schema, plugins: [history()] });
		const command = setReviewRecordingPreferences({ trackFormatting: false, trackMoves: false });
		expect(command(state)).toBe(true);
		expect(state.doc.attrs.trackFormatting).toBe(true);
		command(state, (tr) => {
			state = state.apply(tr);
		});
		expect(state.doc.attrs).toMatchObject({ trackFormatting: false, trackMoves: false });
		let dispatched = false;
		command(state, () => {
			dispatched = true;
		});
		expect(dispatched).toBe(false);
		expect(
			undo(state, (tr) => {
				state = state.apply(tr);
			}),
		).toBe(true);
		expect(state.doc.attrs).toMatchObject({ trackFormatting: true, trackMoves: true });
	});
	it('refuses an unsupported preference without dispatching partial updates', () => {
		const schema = new Schema({
			nodes: { doc: { content: 'text*', attrs: { trackFormatting: { default: true } } }, text: {} },
		});
		let dispatched = false;
		expect(
			setReviewRecordingPreferences({ trackFormatting: false, trackMoves: false })(
				EditorState.create({ schema }),
				() => {
					dispatched = true;
				},
			),
		).toBe(false);
		expect(dispatched).toBe(false);
	});
	it.each([
		['trackFormatting', toggleTrackFormatting],
		['trackMoves', toggleTrackMoves],
	] as const)('toggles %s as an isolated undoable preference', (name, command) => {
		const schema = new Schema({
			nodes: {
				doc: { content: 'text*', attrs: { [name]: { default: true } } },
				text: {},
			},
		});
		let state = EditorState.create({
			doc: schema.node('doc', null, schema.text('Text')),
			plugins: [history()],
		});
		expect(
			command(state, (tr) => {
				state = state.apply(tr);
			}),
		).toBe(true);
		expect(state.doc.attrs[name]).toBe(false);
		expect(
			undo(state, (tr) => {
				state = state.apply(tr);
			}),
		).toBe(true);
		expect(state.doc.attrs[name]).toBe(true);
	});
	it('toggles through a document attribute step and supports undo', () => {
		const schema = new Schema({
			nodes: { doc: { content: 'text*', attrs: { trackChanges: { default: false } } }, text: {} },
		});
		let state = EditorState.create({
			doc: schema.node('doc', null, schema.text('Text')),
			plugins: [history()],
		});
		expect(
			toggleTrackChanges(state, (tr) => {
				state = state.apply(tr);
			}),
		).toBe(true);
		expect(state.doc.attrs.trackChanges).toBe(true);
		expect(state.doc.textContent).toBe('Text');
		expect(
			undo(state, (tr) => {
				state = state.apply(tr);
			}),
		).toBe(true);
		expect(state.doc.attrs.trackChanges).toBe(false);
	});
	it('does not report support for a schema without the recording attribute', () => {
		const schema = new Schema({ nodes: { doc: { content: 'text*' }, text: {} } });
		expect(toggleTrackChanges(EditorState.create({ schema }))).toBe(false);
	});
});
