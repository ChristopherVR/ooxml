/**
 * release-highlight.mjs: a small build-time token colouring for the code samples of a release
 * post (TypeScript and JavaScript, JSON, HTML and shell). It only marks comments, strings,
 * keywords, numbers, tags and commands; the page's CSS colours them. Unknown languages stay plain.
 */
import { escapeHtml } from './release-markdown.mjs';

const KEYWORDS =
	'abstract|as|async|await|break|case|catch|class|const|continue|declare|default|delete|do|else|enum|export|extends|false|finally|for|from|function|if|implements|import|in|instanceof|interface|let|new|null|of|private|protected|public|readonly|return|satisfies|static|super|switch|this|throw|true|try|type|typeof|undefined|var|void|while|yield';
const STRING = String.raw`'(?:\\.|[^'\\\n])*'|"(?:\\.|[^"\\\n])*"`;

const RULES = {
	script: new RegExp(
		[
			String.raw`(?<comment>\/\/[^\n]*|\/\*[\s\S]*?\*\/)`,
			String.raw`(?<string>${STRING}|\x60(?:\\.|[^\x60\\])*\x60)`,
			String.raw`(?<tag></?[A-Za-z][\w.-]*(?=[\s/>]))`,
			String.raw`(?<keyword>\b(?:${KEYWORDS})\b)`,
			String.raw`(?<number>\b\d[\d_]*(?:\.\d+)?\b)`,
		].join('|'),
		'g',
	),
	json: new RegExp(
		[
			String.raw`(?<attr>"(?:\\.|[^"\\\n])*"(?=\s*:))`,
			String.raw`(?<string>"(?:\\.|[^"\\\n])*")`,
			String.raw`(?<keyword>\b(?:true|false|null)\b)`,
			String.raw`(?<number>-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b)`,
		].join('|'),
		'g',
	),
	html: new RegExp(
		[
			String.raw`(?<comment><!--[\s\S]*?-->)`,
			String.raw`(?<tag></?[A-Za-z][\w-]*|/?>)`,
			String.raw`(?<attr>\b[\w:-]+(?==))`,
			String.raw`(?<string>${STRING})`,
		].join('|'),
		'g',
	),
	shell: new RegExp(
		[
			String.raw`(?<comment>(?:^|(?<=\s))#[^\n]*)`,
			String.raw`(?<string>${STRING})`,
			String.raw`(?<command>(?:^|(?<=[|;&]\s*))[\w./@-]+)`,
			String.raw`(?<attr>(?<=\s)--?[\w-]+)`,
		].join('|'),
		'gm',
	),
};

const FAMILY = {
	ts: 'script',
	tsx: 'script',
	typescript: 'script',
	js: 'script',
	jsx: 'script',
	mjs: 'script',
	javascript: 'script',
	vue: 'html',
	svelte: 'html',
	json: 'json',
	jsonc: 'json',
	html: 'html',
	xml: 'html',
	sh: 'shell',
	bash: 'shell',
	shell: 'shell',
	console: 'shell',
	powershell: 'shell',
	ps1: 'shell',
};

/** The label a code block's bar shows for its language. */
export const LANGUAGE_NAMES = {
	ts: 'TypeScript',
	typescript: 'TypeScript',
	tsx: 'TSX',
	js: 'JavaScript',
	javascript: 'JavaScript',
	mjs: 'JavaScript',
	jsx: 'JSX',
	json: 'JSON',
	jsonc: 'JSON',
	html: 'HTML',
	xml: 'XML',
	vue: 'Vue',
	svelte: 'Svelte',
	css: 'CSS',
	sh: 'Shell',
	bash: 'Shell',
	shell: 'Shell',
	console: 'Shell',
	powershell: 'PowerShell',
	ps1: 'PowerShell',
};

/** Code to escaped HTML with `<span class="tok-<kind>">` around the tokens of its language. */
export function highlight(code, lang = '') {
	const rule = RULES[FAMILY[lang]];
	if (!rule) return escapeHtml(code);
	let out = '';
	let last = 0;
	for (const match of code.matchAll(rule)) {
		if (!match[0]) continue;
		const kind = Object.entries(match.groups).find(([, value]) => value !== undefined)?.[0];
		out += escapeHtml(code.slice(last, match.index));
		out += `<span class="tok-${kind}">${escapeHtml(match[0])}</span>`;
		last = match.index + match[0].length;
	}
	return out + escapeHtml(code.slice(last));
}
