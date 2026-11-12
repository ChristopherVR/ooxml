import type {
	BuildingBlockActionId,
	BuildingBlockResult,
	ViewerBuildingBlocksResult,
} from 'pptx-react-viewer';
import React, { useState } from 'react';

/**
 * Host-owned UI for the workflows `useViewerBuildingBlocks` leaves to the host (#285): the
 * building blocks raise a request (`actions.pending`), this panel renders the picker or form, and
 * the matching `actions.*` operation finishes it. Every result is written to a readout so the
 * browser spec can assert request, outcome and failure codes.
 */
const OMML = { 'm:oMath': { 'm:r': { 'm:t': 'x+1' } } };
const OMML_EDITED = { 'm:oMath': { 'm:r': { 'm:t': 'y+2' } } };

function describe(result: BuildingBlockResult): string {
	return result.ok ? 'ok' : result.code;
}

export function HostOwnedWorkflows({
	blocks,
}: {
	blocks: ViewerBuildingBlocksResult;
}): React.ReactElement {
	const { actions, toolbarProps } = blocks;
	const [last, setLast] = useState('none');
	const [url, setUrl] = useState('https://example.test/');
	const [items, setItems] = useState('One,Two,Three');
	const [layoutPath, setLayoutPath] = useState('');
	const [chartKind, setChartKind] = useState<'column' | 'pie'>('pie');
	const report = (name: string, result: BuildingBlockResult) =>
		setLast(`${name}: ${describe(result)}`);
	const file =
		(name: string, run: (picked: File) => Promise<BuildingBlockResult>) =>
		async (event: React.ChangeEvent<HTMLInputElement>) => {
			const picked = event.target.files?.[0];
			if (picked) {
				report(name, await run(picked));
			}
			event.target.value = '';
		};
	const pending = actions.pending?.kind ?? 'none';
	const find = actions.findReplace;
	const layout = actions.layout;
	return (
		<aside aria-label='Host workflows' data-host-workflows style={{ padding: 4, fontSize: 12 }}>
			<div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
				<button type='button' onClick={() => toolbarProps.onOpenImagePicker?.()}>
					Host picture
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenMediaPicker?.()}>
					Host media
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenHyperlinkDialog?.()}>
					Host link
				</button>
				<button type='button' onClick={() => toolbarProps.onAddSmartArt?.()}>
					Host SmartArt
				</button>
				<button type='button' onClick={() => toolbarProps.onAddEquation?.()}>
					Host equation
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenChartDialog?.()}>
					Host chart
				</button>
				<button
					type='button'
					disabled={!toolbarProps.canUndo}
					onClick={() => toolbarProps.onUndo?.()}
				>
					Host undo
				</button>
				<button type='button' onClick={() => actions.cancel()}>
					Cancel request
				</button>
			</div>
			<output aria-label='Host pending request'>{pending}</output>
			<output aria-label='Host last result'>{last}</output>
			<output aria-label='Host action availability'>
				{(Object.keys(actions.availability) as BuildingBlockActionId[])
					.map((id) => {
						const state = actions.availability[id];
						return `${id}=${state.available ? 'yes' : (state.reason ?? 'no')}${state.needsHostUi ? '*' : ''}`;
					})
					.join(' ')}
			</output>
			{actions.pending?.kind === 'image' && (
				<input
					type='file'
					aria-label='Host image file'
					accept='image/*'
					onChange={file('image', actions.insertImage)}
				/>
			)}
			{actions.pending?.kind === 'media' && (
				<input
					type='file'
					aria-label='Host media file'
					accept='audio/*,video/*'
					onChange={file('media', actions.insertMedia)}
				/>
			)}
			{actions.pending?.kind === 'hyperlink' && (
				<div role='group' aria-label='Host hyperlink form'>
					<input aria-label='Host link URL' value={url} onChange={(e) => setUrl(e.target.value)} />
					<button
						type='button'
						onClick={() =>
							report(
								'link',
								actions.hyperlink.apply({
									targetType: 'url',
									url,
									tooltip: '',
									emailAddress: '',
									emailSubject: '',
									slideNumber: 1,
									filePath: '',
									actionVerb: 'nextSlide',
								}),
							)
						}
					>
						Apply link
					</button>
				</div>
			)}
			<button type='button' onClick={() => report('unlink', actions.hyperlink.remove())}>
				Remove link
			</button>
			<output aria-label='Host current link'>{actions.hyperlink.current?.url ?? 'none'}</output>
			{actions.pending?.kind === 'smartArt' && (
				<div role='group' aria-label='Host SmartArt form'>
					<input
						aria-label='Host SmartArt items'
						value={items}
						onChange={(e) => setItems(e.target.value)}
					/>
					<button
						type='button'
						onClick={() =>
							report(
								'smartart',
								actions.smartArt.insert('basic-block-list' as never, items.split(',')),
							)
						}
					>
						Insert SmartArt
					</button>
				</div>
			)}
			{actions.pending?.kind === 'equation' && (
				<button type='button' onClick={() => report('equation', actions.equation.insert(OMML))}>
					Insert equation
				</button>
			)}
			<button
				type='button'
				onClick={() => report('equation', actions.equation.update(OMML_EDITED))}
			>
				Update equation
			</button>
			<output aria-label='Host current equation'>
				{actions.equation.current ? JSON.stringify(actions.equation.current) : 'none'}
			</output>
			{actions.pending?.kind === 'chart' && (
				<div role='group' aria-label='Host chart form'>
					<select
						aria-label='Host chart type'
						value={chartKind}
						onChange={(e) => setChartKind(e.target.value as 'column' | 'pie')}
					>
						<option value='column'>Column</option>
						<option value='pie'>Pie</option>
					</select>
					<button
						type='button'
						onClick={() =>
							report(
								'chart',
								actions.chart.insert(chartKind, {
									title: 'Host chart',
									data: {
										categories: ['North', 'South'],
										series: [{ name: 'Sales', values: [3, 7] }],
									},
								}),
							)
						}
					>
						Insert chart
					</button>
				</div>
			)}
			<div role='group' aria-label='Host find and replace'>
				<input
					aria-label='Host find'
					value={find.query}
					onChange={(e) => find.setQuery(e.target.value)}
				/>
				<input
					aria-label='Host replace'
					value={find.replacement}
					onChange={(e) => find.setReplacement(e.target.value)}
				/>
				<button type='button' onClick={() => find.previous()}>
					Previous match
				</button>
				<button type='button' onClick={() => find.next()}>
					Next match
				</button>
				<button type='button' onClick={() => report('replace', find.replaceCurrent())}>
					Replace one
				</button>
				<button type='button' onClick={() => report('replace', find.replaceAll())}>
					Replace all
				</button>
				<output aria-label='Host match count'>
					{`${find.matches.length} / ${find.currentIndex}`}
				</output>
			</div>
			<div role='group' aria-label='Host layout'>
				<select
					aria-label='Host layout choice'
					value={layoutPath}
					onChange={(e) => setLayoutPath(e.target.value)}
				>
					<option value=''>Choose layout</option>
					{layout.available.map((option) => (
						<option key={option.path} value={option.path}>
							{option.name}
						</option>
					))}
				</select>
				<button
					type='button'
					onClick={async () => report('layout', await layout.apply(layoutPath || 'missing'))}
				>
					Apply layout
				</button>
				<button type='button' onClick={async () => report('layout', await layout.reset())}>
					Reset layout
				</button>
				<output aria-label='Host current layout'>{layout.current?.name ?? 'none'}</output>
			</div>
		</aside>
	);
}
