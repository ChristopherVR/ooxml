import assert from 'node:assert/strict';

/** Test-only HTTP interception. The application always calls its production OneDrive adapter. */
export async function interceptOneDrive(context, fixtures) {
	const files = new Map(
		Object.entries(fixtures).map(([kind, bytes]) => [
			kind,
			{
				id: kind,
				name: `Shared.${kind}`,
				bytes: Buffer.from(bytes),
				revision: 1,
			},
		]),
	);
	const uploads = new Map();
	const writes = [];
	const shares = [];
	/** Graph paths (with their query) the application requested. */
	const requests = [];
	/** Which files each folder lists: the drive root, or the Projects folder inside it. */
	const folders = { root: ['docx', 'xlsx'], projects: ['pptx', 'vsdx'] };
	const metadata = (file) => ({
		id: file.id,
		name: file.name,
		eTag: `"revision-${file.revision}"`,
		parentReference: { driveId: 'test-drive', id: 'test-folder' },
		webUrl: `https://tenant.sharepoint.com/documents/${encodeURIComponent(file.name)}`,
		file: {},
		size: file.bytes.length,
		'@microsoft.graph.downloadUrl': `https://transfer.onedrive-test.invalid/download/${file.id}`,
	});
	const listed = (file, path = '/drive/root:') => ({
		...metadata(file),
		parentReference: { driveId: 'test-drive', id: 'test-folder', path },
		lastModifiedDateTime: '2026-06-01T10:00:00Z',
		lastModifiedBy: { user: { displayName: 'Test User' } },
	});
	const projects = {
		id: 'projects',
		name: 'Projects',
		folder: { childCount: folders.projects.length },
		parentReference: { driveId: 'test-drive', id: 'root', path: '/drive/root:' },
		lastModifiedDateTime: '2026-05-30T09:00:00Z',
	};
	const notes = {
		id: 'notes',
		name: 'Notes.txt',
		file: {},
		size: 12,
		parentReference: { driveId: 'test-drive', id: 'root', path: '/drive/root:' },
	};
	const inFolder = (names, path) => names.map((kind) => listed(files.get(kind), path));
	await context.route('https://graph.microsoft.com/v1.0/**', async (route) => {
		const request = route.request();
		assert.equal(request.headers().authorization, 'Bearer test-only-token');
		const url = new URL(request.url());
		const path = decodeURIComponent(url.pathname.slice('/v1.0'.length));
		requests.push(`${request.method()} ${path}${url.search}`);
		const json = (body, status = 200) => route.fulfill({ status, json: body });
		// The signed-in user's own drive (the PowerPoint demo's edit policy reads it).
		if (path === '/me/drive') return json({ id: 'test-drive' });
		if (path === '/me/drive/root/children')
			return json({
				value: [projects, ...inFolder(folders.root, '/drive/root:'), notes],
			});
		if (path === '/me/drive/recent')
			return json({
				value: [{ ...listed(files.get('docx')), remoteItem: listed(files.get('docx')) }],
			});
		if (path.startsWith('/me/drive/root/search(')) {
			const query = /q='(.*)'\)$/u.exec(path)?.[1]?.toLowerCase() ?? '';
			const hits = [...files.values()].filter((file) => file.name.toLowerCase().includes(query));
			return json({ value: hits.map((file) => listed(file, '/drive/root:/Projects')) });
		}
		if (path.startsWith('/shares/')) {
			const token = path.split('/')[2];
			const shared = new URL(Buffer.from(token.slice(2), 'base64url').toString());
			const file = files.get(shared.pathname.split('/').pop());
			return file ? json(metadata(file)) : json({}, 404);
		}
		const match = /^\/drives\/test-drive\/items\/([^/]+)(.*)$/u.exec(path);
		if (!match) return json({}, 404);
		const id = match[1];
		const action = match[2];
		if (id === 'projects' && action === '/children')
			return json({ value: inFolder(folders.projects, '/drive/root:/Projects') });
		if (id.startsWith('test-folder:') && action.endsWith(':/createUploadSession')) {
			const name = action.slice(1, -':/createUploadSession'.length);
			const copy = { id: `copy-${files.size}`, name, bytes: Buffer.alloc(0), revision: 1 };
			files.set(copy.id, copy);
			uploads.set(copy.id, { file: copy, chunks: [], position: 0 });
			return json({ uploadUrl: `https://transfer.onedrive-test.invalid/upload/${copy.id}` });
		}
		const file = files.get(id);
		if (!file) return json({}, 404);
		if (request.method() === 'GET') return json(metadata(file));
		if (action === '/createLink') {
			shares.push({ id, ...request.postDataJSON() });
			return json({
				link: { webUrl: `https://tenant.sharepoint.com/:x:/shared/${id}?e=test-capability` },
			});
		}
		if (action === '/createUploadSession') {
			if (request.headers()['if-match'] !== metadata(file).eTag) return json({}, 412);
			uploads.set(id, { file, chunks: [], position: 0 });
			return json({ uploadUrl: `https://transfer.onedrive-test.invalid/upload/${id}` });
		}
		return json({}, 404);
	});
	await context.route('https://transfer.onedrive-test.invalid/**', async (route) => {
		const request = route.request();
		assert.equal(request.headers().authorization, undefined);
		const [, action, id] = new URL(request.url()).pathname.split('/');
		if (action === 'download') return route.fulfill({ body: files.get(id).bytes });
		const upload = uploads.get(id);
		assert.ok(upload);
		const range = /^bytes (\d+)-(\d+)\/(\d+)$/u.exec(request.headers()['content-range']);
		assert.ok(range);
		assert.equal(Number(range[1]), upload.position);
		const body = request.postDataBuffer();
		upload.chunks.push(body);
		upload.position += body.length;
		if (upload.position < Number(range[3]))
			return route.fulfill({ status: 202, json: { nextExpectedRanges: [`${upload.position}-`] } });
		upload.file.bytes = Buffer.concat(upload.chunks);
		upload.file.revision++;
		writes.push(id);
		return route.fulfill({ status: 200, json: metadata(upload.file) });
	});
	return { files, writes, shares, requests };
}
