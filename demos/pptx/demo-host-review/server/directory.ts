// The host's own records: who the people are, which review items exist and what each person may
// do with each one. In a real business system this is your user directory and permission store;
// the editor never sees it, it only receives the role the host derives from it.

/** What the host lets an account do with one review item. */
export type Permission = 'owner' | 'review' | 'view' | 'none';

/** The collaboration role the viewer is mounted with for a permission. */
export type ViewerRole = 'owner' | 'collaborator' | 'viewer';

export interface Account {
	id: string;
	name: string;
	title: string;
	color: string;
	/** The host's licence for the AI assistant; an account without it never gets the panel. */
	ai: boolean;
}

export interface Version {
	number: number;
	savedAt: number;
	savedBy: string;
	bytes: Uint8Array;
}

export interface Activity {
	at: number;
	account: string;
	text: string;
}

export interface Review {
	id: string;
	title: string;
	fileName: string;
	source: Uint8Array;
	access: Map<string, Permission>;
	versions: Version[];
	activity: Activity[];
}

export const ACCOUNTS: readonly Account[] = [
	{ id: 'ana', name: 'Ana Ortiz', title: 'Marketing lead', color: '#0f6cbd', ai: true },
	{ id: 'ben', name: 'Ben Okafor', title: 'Legal reviewer', color: '#c239b3', ai: true },
	{ id: 'chen', name: 'Chen Wei', title: 'Finance reviewer', color: '#107c10', ai: false },
	{ id: 'dee', name: 'Dee Patel', title: 'Executive sponsor', color: '#ca5010', ai: false },
];

const MAX_VERSIONS = 10;
const MAX_ACTIVITY = 50;

export const canWrite = (permission: Permission): boolean =>
	permission === 'owner' || permission === 'review';

export function viewerRole(permission: Permission): ViewerRole | null {
	if (permission === 'owner') return 'owner';
	if (permission === 'review') return 'collaborator';
	if (permission === 'view') return 'viewer';
	return null;
}

export function createDirectory(source: Uint8Array) {
	const accounts = new Map(ACCOUNTS.map((account) => [account.id, account]));
	const reviews = new Map<string, Review>([
		[
			'q3-launch',
			{
				id: 'q3-launch',
				title: 'Q3 launch deck',
				fileName: 'q3-launch.pptx',
				source,
				access: new Map<string, Permission>([
					['ana', 'owner'],
					['ben', 'review'],
					['chen', 'review'],
					['dee', 'view'],
				]),
				versions: [],
				activity: [],
			},
		],
	]);

	const permission = (reviewId: string, accountId: string): Permission =>
		reviews.get(reviewId)?.access.get(accountId) ?? 'none';

	const log = (review: Review, account: string, text: string): void => {
		review.activity.unshift({ at: Date.now(), account, text });
		review.activity.length = Math.min(review.activity.length, MAX_ACTIVITY);
	};

	return {
		account: (id: string): Account | undefined => accounts.get(id),
		review: (id: string): Review | undefined => reviews.get(id),
		permission,
		/** AI needs both the account's licence and write access to the item. */
		canUseAi: (reviewId: string, accountId: string): boolean =>
			Boolean(accounts.get(accountId)?.ai) && canWrite(permission(reviewId, accountId)),
		log,
		/** Only the item's owner changes access, and nobody changes the owner's own. */
		setPermission(reviewId: string, by: string, target: string, next: Permission): boolean {
			const review = reviews.get(reviewId);
			if (!review || permission(reviewId, by) !== 'owner') return false;
			if (!accounts.has(target) || target === by || next === 'owner') return false;
			review.access.set(target, next);
			log(review, by, `set ${accounts.get(target)?.name} to ${next}`);
			return true;
		},
		/** A snapshot the owner's editor wrote back. Keeps the newest few. */
		addVersion(reviewId: string, by: string, bytes: Uint8Array): Version | null {
			const review = reviews.get(reviewId);
			if (!review || permission(reviewId, by) !== 'owner') return null;
			const version: Version = {
				number: (review.versions[0]?.number ?? 0) + 1,
				savedAt: Date.now(),
				savedBy: by,
				bytes,
			};
			review.versions.unshift(version);
			review.versions.length = Math.min(review.versions.length, MAX_VERSIONS);
			log(review, by, `saved version ${version.number}`);
			return version;
		},
	};
}

export type Directory = ReturnType<typeof createDirectory>;
