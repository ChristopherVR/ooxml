/**
 * The assistant pane beside the workspace. It is the same chat as the chat page, scoped to the
 * open document: the document is attached when the pane opens, and the suggestions follow its app.
 */
import { $, escape } from './ui.js';
import { icon } from './icons.js';
import { createChat } from './ai/chat.js';

const suggestions = {
	docx: [
		'Summarize this document in five bullet points',
		'Suggest a clearer title and opening paragraph',
		'Turn this document into a short presentation',
	],
	xlsx: [
		'Explain what this workbook calculates',
		'Check the formulas for mistakes',
		'Write a one-page report of these numbers as a Word document',
	],
	pptx: [
		'Summarize this deck',
		'Tighten the wording on every slide',
		'Draft a one-page brief from this deck as a Word document',
	],
	vsdx: ['Explain this diagram step by step', 'List the shapes that have no label'],
	none: [
		'What can you help me with?',
		'Create a presentation about our quarterly results',
		'Make a project tracker workbook',
	],
};

export function mountAssistant({ store, sessions, refresh, open, getActive, getTeams, openChat }) {
	const intro = document.createElement('div');
	intro.className = 'assistant-intro';
	const chat = createChat({
		root: $('assistant-chat'),
		prefix: 'ai',
		deps: { store, sessions, refresh, open, getTeams },
		empty: intro,
		placeholder: 'Ask about your work...',
	});
	async function renderIntro() {
		const active = getActive() && (await store.get(getActive()));
		intro.innerHTML = `<span class="chat-hero-mark" aria-hidden="true">${icon('spark')}</span><h3>${active ? `Working on ${escape(active.name)}` : 'Work across your files'}</h3><p>${
			active
				? 'Ask about this file, have it rewritten, or turn it into something new. Attach more files to compare them.'
				: 'Ask a question, draft something new, or attach files from your library to work across them.'
		}</p><div class="assistant-suggestions">${(suggestions[active?.kind] ?? suggestions.none)
			.map((s) => `<button type="button" data-suggest>${escape(s)}</button>`)
			.join('')}</div>`;
	}
	intro.addEventListener('click', (event) => {
		const b = event.target.closest('[data-suggest]');
		if (b) chat.send(b.textContent);
	});
	$('ai-new').innerHTML = icon('add');
	$('ai-expand').innerHTML = icon('expand');
	$('ai-close').innerHTML = icon('close');
	$('ai-new').onclick = () => {
		chat.reset();
		chat.focus();
	};
	$('ai-expand').onclick = () => openChat(chat.id);
	return {
		/** Called whenever the pane opens: attach the open document and refresh the suggestions. */
		async opened() {
			const active = getActive();
			if (active && !chat.hasAttachments() && !chat.id) await chat.attach([active]);
			await renderIntro();
			chat.render();
			chat.focus();
		},
	};
}
