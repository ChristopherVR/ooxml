import { HostError, getAccounts, getMe, signIn } from './api';
import { h } from './dom';
import { mountReviewPage } from './review-page';

import './styles.css';

const REVIEW_ID = 'q3-launch';
const root = document.getElementById('app');
if (!root) throw new Error('missing #app root');
const app: HTMLElement = root;

/** Demo sign-in. A real host uses its own identity provider; the editor never sees this step. */
async function showSignIn(): Promise<void> {
	const accounts = await getAccounts();
	app.replaceChildren(
		h(
			'main',
			{ class: 'sign-in' },
			h('h1', {}, 'Review hub'),
			h(
				'p',
				{},
				'Sign in as one person here and as another in a second browser or a private window.',
			),
			h(
				'ul',
				{},
				...accounts.map((account) =>
					h(
						'li',
						{},
						h(
							'button',
							{
								type: 'button',
								onclick: () => void signIn(account.id).then(() => location.reload()),
							},
							h('span', { class: 'dot', style: `background:${account.color}` }),
							h('span', { class: 'who' }, account.name, h('small', {}, account.title)),
						),
					),
				),
			),
		),
	);
}

try {
	const me = await getMe();
	await mountReviewPage(app, me, REVIEW_ID);
} catch (error) {
	if (error instanceof HostError && error.status === 401) await showSignIn();
	else {
		app.replaceChildren(h('p', { class: 'notice' }, `Could not open the review: ${String(error)}`));
		throw error;
	}
}
