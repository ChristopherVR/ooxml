import type { BuildingBlockResult, ViewerBuildingBlocksResult } from 'pptx-react-viewer';
import React, { useState } from 'react';

/**
 * Host-owned Set Up Slide Show, Custom Shows, Broadcast and shortcut help for the building
 * blocks: the toolbar controls raise a request (`actions.pending`), this panel renders the form
 * and the matching `actions.*` operation finishes it. Every result is written to a readout so the
 * browser spec can assert request, outcome and failure codes.
 */
const describe = (result: BuildingBlockResult): string => (result.ok ? 'ok' : result.code);

export function HostOwnedShowDialogs({
	blocks,
}: {
	blocks: ViewerBuildingBlocksResult;
}): React.ReactElement {
	const { actions, toolbarProps } = blocks;
	const { setUpSlideShow, customShows, shortcuts, broadcast } = actions;
	const [last, setLast] = useState('none');
	const [name, setName] = useState('Highlights');
	const [to, setTo] = useState('2');
	const report = (label: string, result: BuildingBlockResult) =>
		setLast(`${label}: ${describe(result)}`);
	const kind = actions.pending?.kind ?? 'none';
	const first = customShows.shows[0];
	return (
		<aside
			aria-label='Host show dialogs'
			data-host-show-dialogs
			style={{ padding: 4, fontSize: 12 }}
		>
			<div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
				<button type='button' onClick={() => toolbarProps.onOpenSetUpSlideShow?.()}>
					Host set up show
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenCustomShows?.()}>
					Host custom shows
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenBroadcastDialog?.()}>
					Host broadcast
				</button>
				<button type='button' onClick={() => toolbarProps.onToggleShortcuts()}>
					Host shortcuts
				</button>
			</div>
			<output aria-label='Show dialog request'>{kind}</output>
			<output aria-label='Show dialog result'>{last}</output>
			<output aria-label='Show dialog availability'>
				{(['setUpSlideShow', 'customShows', 'broadcast', 'shortcuts'] as const)
					.map((id) => {
						const state = actions.availability[id];
						return `${id}=${state.available ? 'yes' : (state.reason ?? 'no')}`;
					})
					.join(' ')}
			</output>
			<output aria-label='Broadcast support'>
				{broadcast.supported ? 'supported' : `${broadcast.reason} (${broadcast.requires})`}
			</output>
			{actions.pending?.kind === 'setUpSlideShow' && (
				<div role='group' aria-label='Host set up show form'>
					<output aria-label='Show type'>{setUpSlideShow.current.showType}</output>
					<output aria-label='Show loop'>{String(setUpSlideShow.current.loop)}</output>
					<output aria-label='Show range'>
						{setUpSlideShow.current.slidesMode}:{setUpSlideShow.current.rangeFrom}-
						{setUpSlideShow.current.rangeTo}
					</output>
					<button
						type='button'
						onClick={() => report('setup', setUpSlideShow.apply({ showType: 'kiosk' }))}
					>
						Apply kiosk
					</button>
					<input aria-label='Show range end' value={to} onChange={(e) => setTo(e.target.value)} />
					<button
						type='button'
						onClick={() =>
							report(
								'setup',
								setUpSlideShow.apply({
									showSlidesMode: 'range',
									showSlidesFrom: 1,
									showSlidesTo: Number(to),
								}),
							)
						}
					>
						Apply range
					</button>
				</div>
			)}
			{actions.pending?.kind === 'customShows' && (
				<div role='group' aria-label='Host custom shows form'>
					<input
						aria-label='Custom show name'
						value={name}
						onChange={(e) => setName(e.target.value)}
					/>
					<button
						type='button'
						onClick={() =>
							report(
								'create',
								customShows.create(
									name,
									customShows.slides.slice(0, 2).map((s) => s.rId),
								),
							)
						}
					>
						Create custom show
					</button>
					<button
						type='button'
						disabled={!first}
						onClick={() => first && report('rename', customShows.rename(first.id, `${name} 2`))}
					>
						Rename first show
					</button>
					<button
						type='button'
						disabled={!first}
						onClick={() =>
							first &&
							report(
								'slides',
								customShows.setSlides(
									first.id,
									customShows.slides.slice(0, 1).map((s) => s.rId),
								),
							)
						}
					>
						Keep one slide
					</button>
					<button
						type='button'
						disabled={!first}
						onClick={() => first && report('start', customShows.start(first.id))}
					>
						Start first show
					</button>
					<button
						type='button'
						disabled={!first}
						onClick={() => first && report('remove', customShows.remove(first.id))}
					>
						Delete first show
					</button>
				</div>
			)}
			<output aria-label='Custom shows list'>
				{customShows.shows.map((show) => `${show.name}[${show.slideCount}]`).join(', ') || 'none'}
			</output>
			<output aria-label='Shortcut help'>
				{shortcuts.open ? `open (${shortcuts.catalogue.length} shortcuts)` : 'closed'}
			</output>
			{shortcuts.open && (
				<ul aria-label='Host shortcut list'>
					{shortcuts.catalogue.slice(0, 3).map((entry) => (
						<li key={entry.id}>
							{entry.description}: {entry.keys}
						</li>
					))}
				</ul>
			)}
		</aside>
	);
}
