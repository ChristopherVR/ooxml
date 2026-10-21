// The page's calls to its own host. Nothing here is editor API: the editor only receives what
// `review-session.ts` and `review-page.ts` derive from these records.

export type Permission = 'owner' | 'review' | 'view' | 'none';

export interface Account {
	id: string;
	name: string;
	title: string;
	color: string;
}

export interface Person extends Account {
	permission: Permission;
	online: boolean;
}

export interface ReviewRecord {
	id: string;
	title: string;
	fileName: string;
	permission: Permission;
	ai: boolean;
	people: Person[];
	versions: { number: number; savedAt: number; savedBy: string }[];
	activity: { at: number; account: string; text: string }[];
}

/** What the host grants one mount of the editor. */
export interface Grant {
	ticket: string;
	expiresAt: number;
	permission: Exclude<Permission, 'none'>;
	role: 'owner' | 'collaborator' | 'viewer';
	intent: 'create' | 'join';
	user: { name: string; color: string };
}

export class HostError extends Error {
	constructor(
		readonly status: number,
		message: string,
	) {
		super(message);
	}
}

async function call(path: string, init: RequestInit = {}): Promise<Response> {
	const response = await fetch(path, {
		credentials: 'same-origin',
		...init,
		headers: { 'x-review-hub': '1', ...init.headers },
	});
	if (!response.ok) {
		const body = (await response.json().catch(() => ({}))) as { error?: string };
		throw new HostError(response.status, body.error ?? response.statusText);
	}
	return response;
}

const jsonBody = (value: unknown): RequestInit => ({
	body: JSON.stringify(value),
	headers: { 'Content-Type': 'application/json' },
});

export const getAccounts = async (): Promise<Account[]> => (await call('/api/accounts')).json();
export const getMe = async (): Promise<Account> => (await call('/api/me')).json();
export const signIn = (account: string) =>
	call('/api/session', { method: 'POST', ...jsonBody({ account }) });
export const signOut = () => call('/api/session', { method: 'DELETE' });

export const getReview = async (id: string): Promise<ReviewRecord> =>
	(await call(`/api/reviews/${id}`)).json();

export const getSource = async (id: string): Promise<Uint8Array> =>
	new Uint8Array(await (await call(`/api/reviews/${id}/source`)).arrayBuffer());

export const requestGrant = async (id: string, refresh = false): Promise<Grant> =>
	(
		await call(`/api/reviews/${id}/ticket${refresh ? '?refresh=1' : ''}`, { method: 'POST' })
	).json();

export const saveVersion = (id: string, bytes: Uint8Array) =>
	call(`/api/reviews/${id}/versions`, {
		method: 'POST',
		body: new Blob([bytes as BlobPart]),
		headers: { 'Content-Type': 'application/octet-stream' },
	});

export const setAccess = (id: string, account: string, permission: Permission) =>
	call(`/api/reviews/${id}/access/${account}`, { method: 'PUT', ...jsonBody({ permission }) });
