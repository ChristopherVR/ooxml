/** Turn the existing suite build into a runnable, commit-specific pkg.pr.new template. */
import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

export async function preparePreviewTemplate(root, source) {
	const { repository, sha, prUrl = '' } = source;
	if (!/^[\w.-]+\/[\w.-]+$/.test(repository) || !/^[a-f0-9]{40}$/.test(sha))
		throw new Error('Preview source must include a repository and full commit SHA');
	if (prUrl && !/^https:\/\/github\.com\/[\w.-]+\/[\w.-]+\/pull\/\d+$/.test(prUrl))
		throw new Error('Invalid preview pull request URL');
	const out = join(root, 'suite-dist');
	const template = join(root, 'preview-template');
	const html = await readFile(join(out, 'index.html'), 'utf8');
	const commitUrl = `https://github.com/${repository}/commit/${sha}`;
	await mkdir(join(out, 'scripts'), { recursive: true });
	await cp(new URL('./preview-site.mjs', import.meta.url), join(out, 'scripts/preview-site.mjs'));
	await writeFile(
		join(out, 'package.json'),
		JSON.stringify(
			{
				name: 'ooxml-office-preview',
				version: '0.0.0',
				type: 'module',
				description: `Office suite from ${repository}@${sha.slice(0, 7)}`,
				scripts: { start: 'node scripts/preview-site.mjs 5173 .' },
				stackblitz: { startCommand: 'npm start' },
			},
			null,
			2,
		) + '\n',
	);
	// StackBlitz POSTs every template file to /run. Keep bundled JS and binary assets
	// in an installable preview package instead of exceeding the request body limit.
	await mkdir(template, { recursive: true });
	await writeFile(
		join(template, 'package.json'),
		JSON.stringify(
			{
				name: 'ooxml-office-playground',
				private: true,
				type: 'module',
				scripts: {
					start: 'node node_modules/ooxml-office-preview/scripts/preview-site.mjs 5173 .',
				},
				dependencies: { 'ooxml-office-preview': '*', 'ooxml-core': '*', 'ooxml-ui': '*' },
				stackblitz: { startCommand: 'npm start' },
			},
			null,
			2,
		) + '\n',
	);
	await writeFile(
		join(out, 'README.md'),
		`# OOXML Office preview\n\nSource: [${repository}@${sha.slice(0, 7)}](${commitUrl}).\n` +
			(prUrl ? `\n[Pull request](${prUrl}).\n` : '') +
			'\nRun `npm start` to open the Office suite. Create a document or open your own file.\n' +
			'\nThe start screen opens the bundled standalone apps. Documentation links use the published guides.\n' +
			'\nThe application assets were bundled in CI from this exact source commit, including ' +
			'its core, UI and PowerPoint binding. Editing an installed package does not rebuild ' +
			'the bundled application. The preview dependencies in package.json are available ' +
			'for separate package experiments.\n',
	);
	await writeFile(join(out, 'preview-source.json'), JSON.stringify(source, null, 2) + '\n');
	await cp(join(out, 'README.md'), join(template, 'README.md'));
	await cp(join(out, 'preview-source.json'), join(template, 'preview-source.json'));
	await writeFile(
		join(out, 'index.html'),
		html.replace(
			'<title>OOXML Office</title>',
			`<meta name="ooxml-preview" content="true" />\n<title>OOXML Office preview: ${sha.slice(0, 7)}</title>`,
		),
	);
	// The preview bundles the apps, not the viewers' separate documentation and demo builds.
	// Keep the local documentation hub, but send its viewer links to the published sites.
	const docsPath = join(out, 'docs/index.html');
	const docs = await readFile(docsPath, 'utf8');
	await writeFile(
		docsPath,
		docs.replace(
			/href="\.\.\/(docx|xlsx|pptx|visio|teams)([\/#][^"]*|)"/g,
			'href="https://christophervr.github.io/ooxml/$1$2"',
		),
	);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href)
	await preparePreviewTemplate(process.cwd(), {
		repository: process.env.PREVIEW_REPOSITORY ?? '',
		sha: process.env.PREVIEW_SHA ?? '',
		prUrl: process.env.PREVIEW_PR_URL ?? '',
	});
