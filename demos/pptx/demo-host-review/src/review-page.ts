// The host's review screen: its own header and side panel around the PowerPoint editor. The page
// owns the session's lifetime; the editor is mounted with the role and readiness the host grants
// and is remounted, on a fresh session, whenever the host changes that grant.
import { createPptxViewer } from 'pptx-vanilla-viewer';
import type { CollaborationConfig, PptxAiConfig, PptxViewerInstance } from 'pptx-vanilla-viewer';

import { HostError, getReview, getSource, saveVersion, setAccess, signOut } from './api';
import type { Account, Permission, ReviewRecord } from './api';
import { h, time } from './dom';
import { openReviewSession } from './review-session';
import type { ReviewSession } from './review-session';

const POLL_MS = 4000;
const PERMISSION_TEXT: Record<Permission, string> = {
	owner: 'Owner: edits, manages access and saves versions',
	review: 'Reviewer: edits live with the others',
	view: 'Viewer: follows along, read-only',
	none: 'No access',
};

function aiConfig(reviewId: string): PptxAiConfig {
	return {
		// The key stays on the host: the route checks the account's licence and access first.
		connection: {
			kind: 'endpoint',
			api: `/api/ai/chat?review=${encodeURIComponent(reviewId)}`,
			headers: { 'x-review-hub': '1' },
			credentials: 'same-origin',
		},
		// Every proposed change waits for the reviewer to accept it.
		writePolicy: 'stage',
	};
}

export async function mountReviewPage(
	root: HTMLElement,
	me: Account,
	reviewId: string,
): Promise<void> {
	let record = await getReview(reviewId);
	const source = await getSource(reviewId);
	const status = h('span', { class: 'status', role: 'status', 'aria-label': 'Session status' });
	const stage = h('main', { class: 'stage', 'aria-label': 'Presentation editor' });
	const side = h('aside', { class: 'side', 'aria-label': 'Review details' });
	root.replaceChildren(
		h(
			'header',
			{ class: 'topbar' },
			h('strong', { class: 'brand' }, 'Review hub'),
			h('span', { class: 'title' }, record.title),
			status,
			h(
				'span',
				{ class: 'me' },
				h('span', { class: 'dot', style: `background:${me.color}` }),
				me.name,
			),
			h(
				'button',
				{ type: 'button', onclick: () => void signOut().then(() => location.reload()) },
				'Sign out',
			),
		),
		h('div', { class: 'body' }, stage, side),
	);

	let session: ReviewSession | undefined;
	let viewer: PptxViewerInstance | undefined;
	let unsubscribe: (() => void) | undefined;

	const renderStatus = (): void => {
		const snapshot = session?.external.getSnapshot();
		const readOnly = session?.grant.role === 'viewer';
		status.textContent = !snapshot
			? 'Opening'
			: snapshot.status !== 'connected'
				? snapshot.synced
					? 'Offline: changes are kept and sent on reconnect'
					: 'Connecting'
				: !snapshot.synced
					? 'Loading the shared deck'
					: readOnly
						? 'Live, read-only'
						: 'Live';
		status.dataset.state = snapshot?.synced && snapshot.status === 'connected' ? 'live' : 'waiting';
	};

	const stop = (): void => {
		unsubscribe?.();
		viewer?.destroy();
		session?.destroy();
		viewer = undefined;
		session = undefined;
	};

	const start = async (): Promise<void> => {
		session = await openReviewSession(reviewId, () => void restart());
		const { grant, external } = session;
		const collaboration: CollaborationConfig = {
			roomId: reviewId,
			serverUrl: '',
			userName: grant.user.name,
			userColor: grant.user.color,
			role: grant.role,
			sessionIntent: grant.intent,
			externalSession: external,
		};
		if (grant.role === 'owner') {
			// One elected writer turns the shared deck into saved versions in the host's store.
			collaboration.onWriteBack = (bytes) => void saveVersion(reviewId, bytes).then(refresh);
			collaboration.writeBackDebounceMs = 8000;
		}
		viewer = createPptxViewer(stage, {
			source: source.slice(),
			fileName: record.fileName,
			editable: grant.role !== 'viewer',
			collaboration,
			...(record.ai ? { ai: aiConfig(reviewId) } : {}),
		});
		unsubscribe = external.subscribe(renderStatus);
		renderStatus();
	};

	let restarting: Promise<void> | undefined;
	const restart = (): Promise<void> =>
		(restarting ??= (async () => {
			stop();
			renderStatus();
			try {
				record = await getReview(reviewId);
				renderSide();
				await start();
			} catch (error) {
				if (error instanceof HostError && error.status === 404) {
					stage.replaceChildren(
						h('p', { class: 'notice' }, 'Your access to this review was removed.'),
					);
					status.textContent = 'No access';
				} else throw error;
			} finally {
				restarting = undefined;
			}
		})());

	const renderSide = (): void => {
		const owner = record.permission === 'owner';
		const name = (id: string): string => record.people.find((p) => p.id === id)?.name ?? id;
		const accessControl = (person: ReviewRecord['people'][number]): Node => {
			if (!owner || person.id === me.id) return h('span', { class: 'perm' }, person.permission);
			const select = h(
				'select',
				{
					'aria-label': `Access for ${person.name}`,
					onchange: (event) => {
						const next = (event.target as HTMLSelectElement).value as Permission;
						void setAccess(reviewId, person.id, next).then(refresh);
					},
				},
				...(['review', 'view', 'none'] as const).map((value) =>
					h('option', value === person.permission ? { value, selected: '' } : { value }, value),
				),
			);
			return select;
		};
		side.replaceChildren(
			h(
				'section',
				{},
				h('h2', {}, 'Your access'),
				h('p', {}, PERMISSION_TEXT[record.permission]),
				record.ai
					? h(
							'p',
							{ class: 'hint' },
							'The AI assistant (sparkles icon) proposes changes you accept or reject.',
						)
					: null,
			),
			h(
				'section',
				{},
				h('h2', {}, 'People'),
				h(
					'ul',
					{ class: 'people' },
					...record.people
						.filter((person) => person.permission !== 'none' || owner)
						.map((person) =>
							h(
								'li',
								{},
								h('span', {
									class: person.online ? 'dot online' : 'dot',
									style: `background:${person.color}`,
									title: person.online ? 'Online' : 'Away',
								}),
								h('span', { class: 'who' }, person.name, h('small', {}, person.title)),
								accessControl(person),
							),
						),
				),
			),
			h(
				'section',
				{},
				h('h2', {}, 'Versions'),
				record.versions.length === 0
					? h(
							'p',
							{ class: 'hint' },
							"Saved automatically from the owner's editor after changes settle.",
						)
					: h(
							'ul',
							{ class: 'list' },
							...record.versions.map((version) =>
								h(
									'li',
									{},
									h(
										'a',
										{ href: `/api/reviews/${reviewId}/versions/${version.number}` },
										`Version ${version.number}`,
									),
									` ${time(version.savedAt)}, ${name(version.savedBy)}`,
								),
							),
						),
			),
			h(
				'section',
				{},
				h('h2', {}, 'Activity'),
				h(
					'ul',
					{ class: 'list' },
					...record.activity
						.slice(0, 12)
						.map((entry) => h('li', {}, `${time(entry.at)} ${name(entry.account)} ${entry.text}`)),
				),
			),
		);
	};

	async function refresh(): Promise<void> {
		try {
			const next = await getReview(reviewId);
			const changed = next.permission !== record.permission || next.ai !== record.ai;
			record = next;
			renderSide();
			// Access granted again after it was removed, or changed without the relay noticing yet.
			if (changed || !session) await restart();
		} catch (error) {
			const lost = error instanceof HostError && (error.status === 404 || error.status === 401);
			if (lost && session) await restart();
		}
	}

	renderSide();
	await start();
	const poll = setInterval(() => void refresh(), POLL_MS);
	window.addEventListener('pagehide', () => {
		clearInterval(poll);
		stop();
	});
}
