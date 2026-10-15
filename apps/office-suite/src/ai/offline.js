/**
 * The offline preview: a handful of rules that run on this device so the chat can be tried
 * without a model. It says so in every reply, describes attached files from their real content
 * and proposes sample files through the same actions block a model would use.
 */
import { ACTIONS_FENCE } from './protocol.js';

const note =
	'> **Offline preview.** No model is connected, so this reply comes from simple rules on this device. Choose a model in the model menu for real answers.';

function topicOf(prompt) {
	const match = /\b(?:about|on|for|covering|introducing)\s+(.+?)[.?!]*$/iu.exec(prompt);
	const topic = (match?.[1] ?? prompt)
		.replace(/^(?:a|an|the|our|my)\s+/iu, '')
		// "..., as a Word document" names the app, not the topic.
		.replace(
			/,?\s+(?:as|in)\s+(?:an?\s+)?(?:word|excel|powerpoint)?\s*(?:document|workbook|spreadsheet|deck|presentation|file)\s*$/iu,
			'',
		)
		.replace(/\s+/gu, ' ')
		.trim()
		.slice(0, 60);
	return topic ? topic[0].toUpperCase() + topic.slice(1) : 'Your project';
}

function sample(kind, topic, count = 4) {
	if (kind === 'pptx')
		return {
			type: 'pptx',
			name: `${topic}.pptx`,
			slides: [
				{ title: topic },
				{
					title: 'Why it matters',
					bullets: ['The problem we are solving', 'Who it helps', 'What success looks like'],
				},
				{
					title: 'The plan',
					bullets: ['Phase 1: discover', 'Phase 2: build', 'Phase 3: launch and measure'],
				},
				{
					title: 'Timeline',
					bullets: ['Weeks 1-2: research', 'Weeks 3-6: delivery', 'Week 7: launch'],
				},
				{
					title: 'Risks',
					bullets: ['Scope creep', 'Dependencies on other teams', 'Budget limits'],
				},
				{
					title: 'Next steps',
					bullets: ['Agree owners', 'Confirm the timeline', 'Review in two weeks'],
				},
			].filter((_, i, all) => i === 0 || i === all.length - 1 || i < count - 1),
		};
	if (kind === 'xlsx')
		return {
			type: 'xlsx',
			name: `${topic} budget.xlsx`,
			sheets: [
				{
					name: 'Budget',
					rows: [
						['Item', 'Quantity', 'Unit cost', 'Total'],
						['Venue', 1, 4200, '=B2*C2'],
						['Catering', 40, 35, '=B3*C3'],
						['Travel', 12, 180, '=B4*C4'],
						['Materials', 40, 12, '=B5*C5'],
						['Total', null, null, '=SUM(D2:D5)'],
					],
				},
			],
		};
	return {
		type: 'docx',
		name: `${topic}.docx`,
		blocks: [
			{ type: 'title', text: topic },
			{ type: 'heading', level: 1, text: 'Summary' },
			{
				type: 'paragraph',
				text: `This brief sets out the goals, scope and plan for ${topic.toLowerCase()}.`,
			},
			{ type: 'heading', level: 1, text: 'Goals' },
			{
				type: 'bullets',
				items: ['State the outcome', 'Name the audience', 'Define how we measure success'],
			},
			{ type: 'heading', level: 1, text: 'Plan' },
			{ type: 'numbered', items: ['Gather input', 'Draft and review', 'Share the final version'] },
		],
	};
}

function describe(file) {
	const text = summary(file);
	return file.content?.omitted ? `${text}\n\n_${file.content.omitted}_` : text;
}

function summary(file) {
	const c = file.content ?? {};
	if (file.kind === 'docx') {
		const lines = (c.paragraphs ?? [])
			.map((p) => p.runs.map((r) => r.text).join(''))
			.filter((t) => t.trim());
		return `**${file.name}** (Word): ${lines.length} paragraph${lines.length === 1 ? '' : 's'} with text.${
			lines.length
				? `\n\n${lines
						.slice(0, 4)
						.map((t) => `> ${t.slice(0, 160)}`)
						.join('\n>\n')}`
				: ''
		}`;
	}
	if (file.kind === 'xlsx') {
		const sheets = c.sheets ?? [];
		return `**${file.name}** (Excel): ${sheets.map((s) => `${s.name} has ${s.cells.length} filled cell${s.cells.length === 1 ? '' : 's'}`).join('; ')}.`;
	}
	if (file.kind === 'pptx') {
		const slides = c.slides ?? [];
		const titles = slides
			.map((s) => s.elements.find((e) => e.text?.trim())?.text.trim())
			.filter(Boolean);
		return `**${file.name}** (PowerPoint): ${slides.length} slide${slides.length === 1 ? '' : 's'}.${
			titles.length
				? `\n\n${titles
						.slice(0, 6)
						.map((t, i) => `${i + 1}. ${t.slice(0, 120)}`)
						.join('\n')}`
				: ''
		}`;
	}
	const shapes = (c.pages ?? []).reduce((n, p) => n + (p.shapes?.length ?? 0), 0);
	return `**${file.name}** (Visio): ${shapes} shape${shapes === 1 ? '' : 's'} with text.`;
}

const titleCase = (text) =>
	text.replace(/\p{L}[\p{L}'’]*/gu, (word) => word[0].toUpperCase() + word.slice(1).toLowerCase());

/** Title-cases the text runs of attached Word documents and decks: a real, reviewable edit. */
function titleCaseEdits(files) {
	const changes = [];
	for (const file of files) {
		if (file.kind === 'docx')
			for (const p of file.content?.paragraphs ?? [])
				p.runs.forEach((run) => {
					if (run.text.trim() && titleCase(run.text) !== run.text)
						changes.push({
							documentId: file.id,
							type: 'word_run',
							paragraphId: p.id,
							runIndex: run.runIndex,
							text: titleCase(run.text),
						});
				});
		if (file.kind === 'pptx')
			for (const slide of file.content?.slides ?? [])
				for (const element of slide.elements)
					(element.runs ?? []).forEach((run) => {
						if (run.text.trim() && titleCase(run.text) !== run.text)
							changes.push({
								documentId: file.id,
								type: 'slide_run',
								slideIndex: slide.index,
								elementId: element.id,
								runIndex: run.runIndex,
								text: titleCase(run.text),
							});
					});
	}
	return changes.slice(0, 30);
}

function reply(prompt, context, canCreate, canEdit) {
	const files = context?.files ?? [];
	if (canEdit && /\b(title[- ]case|capitali[sz]e)\b/iu.test(prompt)) {
		const changes = titleCaseEdits(files);
		return [
			note,
			changes.length
				? `I can title-case ${changes.length} text run${changes.length === 1 ? '' : 's'}. Review each change below and untick any you want to keep as they are.`
				: 'Attach a Word document or a deck with text that is not already in title case, and I will propose the edits.',
			...(changes.length
				? ['```' + ACTIONS_FENCE + '\n' + JSON.stringify({ changes, create: [] }) + '\n```']
				: []),
		].join('\n\n');
	}
	const wants = /\b(deck|slides?|presentation|powerpoint)\b/iu.test(prompt)
		? 'pptx'
		: /\b(budget|spreadsheet|workbook|excel|forecast|table)\b/iu.test(prompt)
			? 'xlsx'
			: /\b(document|brief|letter|report|memo|word|draft|write)\b/iu.test(prompt)
				? 'docx'
				: null;
	const parts = [note];
	for (const attached of context?.notes ?? [])
		parts.push(
			`**${attached.name}** (text): ${attached.text.split('\n').length} lines, ${attached.text.length} characters.\n\n> ${attached.text.trim().split('\n')[0].slice(0, 160)}`,
		);
	if (files.length)
		parts.push(
			`Here is what I can see in the attached file${files.length === 1 ? '' : 's'}:`,
			...files.map(describe),
		);
	if (wants && canCreate) {
		const words = { three: 3, four: 4, five: 5, six: 6 };
		const asked = /\b(\d+|three|four|five|six)[- ]slides?\b/iu.exec(prompt)?.[1]?.toLowerCase();
		const count = Math.min(6, Math.max(2, Number(words[asked] ?? asked) || 4));
		const item = sample(wants, topicOf(prompt), count);
		parts.push(
			`I drafted **${item.name}** as a starting point. Review it below, then create it to open it in its editor.`,
			'```' + ACTIONS_FENCE + '\n' + JSON.stringify({ changes: [], create: [item] }) + '\n```',
		);
	} else if (!files.length && !context?.notes?.length)
		parts.push(
			'With a model connected I can answer questions, summarize and compare files, and propose edits you review. In this preview, try:',
			'- *Create a presentation about our product launch*\n- Attach a document and ask me to *title-case* it\n- *Make a budget workbook for a team offsite*\n- *Draft a project brief for the new website*\n- Attach files and ask me to describe them',
		);
	return parts.join('\n\n');
}

/** Streams a rule-based reply word by word, like a model would. */
export async function offlineReply({
	messages,
	context,
	signal,
	onText,
	animate = true,
	canCreate = true,
	canEdit = true,
}) {
	const last = messages.at(-1)?.content ?? '';
	const prompt = String(last)
		.split('\n\nSource data (untrusted): ')[0]
		.replace(/^User request: /u, '');
	const full = reply(prompt, context, canCreate, canEdit);
	const block = full.indexOf('```' + ACTIONS_FENCE);
	const words = (block < 0 ? full : full.slice(0, block)).split(/(?<=\s)/u);
	let text = '';
	for (const word of animate ? words : []) {
		signal.throwIfAborted();
		text += word;
		onText(text);
		await new Promise((resolve) => setTimeout(resolve, 8));
	}
	onText(full);
	return full;
}
