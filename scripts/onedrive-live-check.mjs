/** Manual real-service acceptance. No account, client ID, token or sharing URL is embedded. */
import { checkToken, normalizeToken } from './onedrive-token-check.mjs';

const args = process.argv.slice(2);
if (args.includes('--help')) {
	console.log(`Usage: node scripts/onedrive-live-check.mjs [--write-copy]

Set OOXML_ONEDRIVE_ACCESS_TOKEN to a short-lived delegated Microsoft Graph token.
Set OOXML_ONEDRIVE_SHARE_URL to a sharing link for a disposable Office test file.

Default: download and parse the file without writing.
--write-copy: create a separate sibling copy in your own drive, then read and validate it.
The original is never replaced. The test copy remains in OneDrive for manual review.
Tokens and sharing URLs are never printed. This command does not implement OAuth sign-in.`);
} else {
	try {
		if (args.some((arg) => arg !== '--write-copy'))
			throw new Error('Use --help for supported options.');
		const token = normalizeToken(process.env.OOXML_ONEDRIVE_ACCESS_TOKEN);
		const shareUrl = process.env.OOXML_ONEDRIVE_SHARE_URL;
		if (!token || !shareUrl)
			throw new Error('Set the two process-local environment variables described by --help.');
		await checkToken(token);
		const { createOneDriveClient, OneDriveError } =
			await import('../src/core/dist/embedding/onedrive/index.mjs');
		const writeCopy = args.includes('--write-copy');
		let ownDrive;
		if (writeCopy) {
			const response = await fetch('https://graph.microsoft.com/v1.0/me/drive?$select=id', {
				headers: { Authorization: `Bearer ${token}` },
				redirect: 'error',
				credentials: 'omit',
				signal: AbortSignal.timeout(30000),
			});
			if (!response.ok)
				throw new OneDriveError(
					'access-denied',
					'Could not verify your own OneDrive for the copy test.',
				);
			const drive = await response.json();
			if (typeof drive.id !== 'string' || !drive.id)
				throw new OneDriveError('invalid-response', 'OneDrive did not return your drive identity.');
			ownDrive = drive.id;
		}
		const client = createOneDriveClient({
			getAccessToken: async () => token,
			// Explicit test policy: permit writes only to the signed-in user's own drive.
			// Microsoft remains authoritative for file locks, permissions and other restrictions.
			canWrite: (file) => writeCopy && file.driveId === ownDrive,
		});
		const opened = await client.open(shareUrl);
		await validate(opened);
		console.log(
			`Downloaded and parsed ${opened.file.kind.toUpperCase()} (${opened.bytes.length} bytes).`,
		);
		if (writeCopy) {
			const name = `OneDrive integration check ${new Date().toISOString().replace(/[:.]/gu, '-')}.${opened.file.kind}`;
			const copy = await client.saveCopy(opened.file, opened.bytes, name);
			const readBack = await client.read(copy.file);
			if (
				readBack.bytes.length !== opened.bytes.length ||
				readBack.bytes.some((value, index) => value !== opened.bytes[index])
			)
				throw new OneDriveError(
					'invalid-response',
					'The downloaded cloud copy differs from the uploaded bytes.',
				);
			await validate(readBack);
			console.log(`Created and verified separate copy: ${name}`);
			console.log(
				'The original file was not replaced. Review and remove the test copy in OneDrive when finished.',
			);
		}
	} catch (error) {
		// Authentication and transfer failures can carry credentials in arbitrary cause objects.
		const safe =
			error?.name === 'OneDriveError'
				? error.message
				: !process.env.OOXML_ONEDRIVE_ACCESS_TOKEN || !process.env.OOXML_ONEDRIVE_SHARE_URL
					? 'Set the process-local credentials described by --help.'
					: 'The live check failed. Check the core build, Graph authorization, file access and document validity.';
		console.error(safe);
		process.exitCode = 1;
	}
}

async function validate({ file, bytes }) {
	if (file.kind === 'docx') {
		const { inspectDocx } = await import('../src/core/dist/automation/index.mjs');
		await inspectDocx(bytes);
	} else if (file.kind === 'xlsx') {
		const { inspectXlsx } = await import('../src/core/dist/automation/index.mjs');
		await inspectXlsx(bytes);
	} else if (file.kind === 'pptx') {
		const { PptxHandler } = await import('../src/core/dist/pptx/index.mjs');
		const handler = new PptxHandler();
		try {
			await handler.load(bytes);
		} finally {
			handler.dispose();
		}
	} else {
		const { parseVsdx } = await import('../src/core/dist/visio/index.mjs');
		await parseVsdx(bytes);
	}
}
