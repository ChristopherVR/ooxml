/**
 * Shared by the OneDrive live scripts: clean up a pasted access token and, before any measurement,
 * check it against Microsoft Graph. A rejection is explained without ever printing the token: an
 * authorization code pasted by mistake, an expired token, a token for another resource or without
 * file scopes, and Microsoft's own error description.
 */

const GRAPH_AUDIENCES = new Set([
	'https://graph.microsoft.com',
	'https://graph.microsoft.com/',
	'00000003-0000-0000-c000-000000000000',
]);

/** Trim a pasted token and drop a leading `Bearer ` and surrounding quotes. */
export function normalizeToken(value) {
	let token = String(value ?? '').trim();
	token = token.replace(/^(['"])(.*)\1$/su, '$2').trim();
	token = token.replace(/^Bearer\s+/iu, '').trim();
	return token;
}

function decodeJwtPayload(token) {
	const parts = token.split('.');
	if (parts.length !== 3) return undefined;
	try {
		return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
	} catch {
		return undefined;
	}
}

/**
 * What the token looks like, from its shape alone, as safe facts (never the token). Work and school
 * tokens are JWTs whose claims say the audience, scopes and expiry; personal Microsoft account
 * tokens are opaque. An authorization code from a redirect URL is not a token at all.
 */
export function describeToken(token, nowSeconds = Math.floor(Date.now() / 1000)) {
	const notes = [];
	if (/\s/u.test(token)) notes.push('it contains spaces or line breaks; paste the token only');
	if (/^M\.[A-Za-z0-9_]/u.test(token) || /^0\.A[A-Za-z0-9_-]/u.test(token))
		notes.push(
			'it looks like an authorization code (the code= value of a sign-in redirect), not an access token; copy the token from Graph Explorer’s "Access token" tab instead',
		);
	const claims = decodeJwtPayload(token);
	if (claims) {
		if (claims.aud && !GRAPH_AUDIENCES.has(String(claims.aud)))
			notes.push(
				`it is for another resource (audience ${String(claims.aud)}), not Microsoft Graph`,
			);
		if (typeof claims.exp === 'number' && claims.exp <= nowSeconds)
			notes.push(`it expired ${Math.round((nowSeconds - claims.exp) / 60)} minutes ago`);
		const scopes = String(claims.scp ?? '')
			.split(' ')
			.filter(Boolean);
		if (scopes.length && !scopes.some((scope) => /^Files\./u.test(scope)))
			notes.push(`it has no Files.* scope (scopes: ${scopes.join(' ')})`);
		if (!claims.scp && claims.roles)
			notes.push('it is an application token; use a delegated (signed-in user) token');
	} else if (!notes.length && token.length < 100) {
		notes.push(`it is only ${token.length} characters long; an access token is usually over 1000`);
	}
	return { kind: claims ? 'jwt' : 'opaque', notes };
}

/**
 * Ask Graph who the token belongs to. Resolves on success; throws an Error whose message is safe to
 * print (it never contains the token) when Graph rejects it.
 */
export async function checkToken(token, fetcher = globalThis.fetch) {
	let response;
	try {
		response = await fetcher('https://graph.microsoft.com/v1.0/me/drive?$select=id', {
			headers: { Authorization: `Bearer ${token}` },
			credentials: 'omit',
			redirect: 'error',
			signal: AbortSignal.timeout(30000),
		});
	} catch {
		throw new Error('Could not reach Microsoft Graph. Check the network connection.');
	}
	if (response.ok) return;
	const lines = [`Microsoft Graph rejected the token (HTTP ${response.status}).`];
	try {
		const body = await response.json();
		const code = body?.error?.code;
		const message = body?.error?.message;
		if (code || message)
			lines.push(`Graph says: ${[code, message].filter(Boolean).join(': ').slice(0, 300)}`);
	} catch {
		// No JSON body.
	}
	const challenge = response.headers.get('www-authenticate');
	const description = challenge && /error_description="([^"]*)"/u.exec(challenge)?.[1];
	if (description) lines.push(`Sign-in service says: ${description.slice(0, 300)}`);
	const { notes } = describeToken(token);
	if (notes.length) lines.push(...notes.map((note) => `The value you set: ${note}.`));
	else if (response.status === 401)
		lines.push(
			'The token is not accepted: it may have expired (Graph Explorer tokens last about an hour) or been copied incompletely. Copy a fresh one from Graph Explorer’s "Access token" tab.',
		);
	if (response.status === 403)
		lines.push(
			'The token is valid but lacks access: consent to Files.ReadWrite (or Files.Read) in Graph Explorer.',
		);
	throw new Error(lines.join('\n'));
}
