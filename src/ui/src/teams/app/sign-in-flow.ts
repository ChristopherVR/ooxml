// The session lifecycle of a Teams server that requires accounts, kept apart from the element so it
// can be tested without a DOM: find out whether the server needs a sign-in (`/health`), exchange
// the credential for a session, notice when the session stops working (it expires, or the server
// rejects it) and sign out. The session lives in memory only: it is secret and short-lived, so a
// reload signs in again. A server without accounts (the zero-config dev default) never gates the
// app, and a host that passes `config.sessionToken` itself manages its own sessions.
import {
	type SessionCheck,
	type TeamsServerConfig,
	type TeamsSession,
	TeamsSignInError,
	checkTeamsSession,
	probeTeamsAccounts,
	requestTeamsSession,
	serverOrigin,
	signOutTeamsSession,
} from 'ooxml-core/teams';
import { SIGN_IN_STRINGS, signInErrorText } from './sign-in-strings';

export type AuthPhase = 'open' | 'checking' | 'required';

export interface SignInHost {
	/** The sign-in state changed: re-render and, if the client must (re)start, do it. */
	changed(): void;
	/** Tell the embedding page (`teams-signed-in`, `teams-signed-out`). */
	emit(name: string, detail: unknown): void;
}

/** How long a lost connection is tolerated before asking the server whether the session is dead. */
const PROBE_AFTER_MS = 4000;
const PROBE_AGAIN_MS = 15_000;
/** `setTimeout` takes 32 bit delays. */
const MAX_TIMER_MS = 2 ** 31 - 1;

export class SignInFlow {
	phase: AuthPhase = 'open';
	session: TeamsSession | null = null;
	/** Kept across a re-prompt so the person only retypes the password. */
	userId = '';
	notice = '';
	error = '';
	busy = false;
	retryAfterSec = 0;
	private readonly accounts = new Map<string, boolean>();
	private probing = '';
	private expiryTimer: ReturnType<typeof setTimeout> | undefined;
	private checkTimer: ReturnType<typeof setTimeout> | undefined;
	private checking = false;
	private generation = 0;

	constructor(private readonly host: SignInHost) {}

	private key(config: TeamsServerConfig): string {
		return `${serverOrigin(config) ?? ''}|${config.token ?? ''}`;
	}

	/**
	 * Whether the client must wait: true while the server is being asked and while the sign-in
	 * screen is showing. Call it from every (re)start with the settings-derived config.
	 */
	gate(config: TeamsServerConfig): boolean {
		if (config.mode !== 'server' || config.sessionToken || !serverOrigin(config)) {
			this.phase = 'open';
			return false;
		}
		if (this.session) {
			this.phase = 'open';
			return false;
		}
		const key = this.key(config);
		const known = this.accounts.get(key);
		if (known === false) {
			this.phase = 'open';
			return false;
		}
		if (known === true) {
			this.phase = 'required';
			return true;
		}
		this.phase = 'checking';
		if (this.probing !== key) {
			this.probing = key;
			const generation = this.generation;
			void probeTeamsAccounts(config).then((accounts) => {
				if (generation !== this.generation || this.probing !== key) return;
				this.probing = '';
				// Unreachable or not a Teams server: carry on exactly as before, and ask again later.
				if (accounts !== null) this.accounts.set(key, accounts);
				this.phase = accounts ? 'required' : 'open';
				this.host.changed();
			});
		}
		return true;
	}

	/** The config to start the client with: the session, when there is one, rides along. */
	apply(config: TeamsServerConfig): TeamsServerConfig {
		return this.session && !config.sessionToken
			? { ...config, sessionToken: this.session.token }
			: config;
	}

	async submit(config: TeamsServerConfig, userId: string, credential: string): Promise<void> {
		if (this.busy) return;
		const id = userId.trim();
		this.userId = id;
		this.error = '';
		this.retryAfterSec = 0;
		this.busy = true;
		this.host.changed();
		try {
			const session = await requestTeamsSession(config, { userId: id, credential });
			this.session = session;
			this.userId = session.userId;
			this.notice = '';
			this.phase = 'open';
			this.armExpiry(session);
			this.host.emit('teams-signed-in', { userId: session.userId, expiresAt: session.expiresAt });
		} catch (error) {
			const failure = error instanceof TeamsSignInError ? error : null;
			this.error = signInErrorText(failure?.reason ?? 'server', failure?.retryAfterSec);
			this.retryAfterSec = failure?.retryAfterSec ?? 0;
		} finally {
			this.busy = false;
			this.host.changed();
		}
	}

	/** The session no longer works: back to the sign-in screen, keeping the user id. */
	expire(message: string): void {
		if (!this.session) return;
		this.drop();
		this.notice = message;
		this.host.emit('teams-session-ended', { userId: this.userId });
		this.host.changed();
	}

	async signOut(config: TeamsServerConfig): Promise<void> {
		const session = this.session;
		if (!session) return;
		this.drop();
		this.userId = '';
		this.notice = SIGN_IN_STRINGS.signedOut;
		this.host.emit('teams-signed-out', { userId: session.userId });
		this.host.changed();
		// Ending it on the server is best effort: the local session is gone either way.
		await signOutTeamsSession(config, session.token);
	}

	/** Feed the connection status after every render; a session that stays unusable is checked. */
	watchConnection(connected: boolean, config: TeamsServerConfig): void {
		if (!this.session) return;
		if (connected) {
			clearTimeout(this.checkTimer);
			this.checkTimer = undefined;
			return;
		}
		if (this.checkTimer !== undefined || this.checking) return;
		this.scheduleCheck(config, PROBE_AFTER_MS);
	}

	private scheduleCheck(config: TeamsServerConfig, delay: number): void {
		const session = this.session;
		if (!session) return;
		this.checkTimer = setTimeout(() => {
			this.checkTimer = undefined;
			this.checking = true;
			void checkTeamsSession(config, session.token).then((result: SessionCheck) => {
				this.checking = false;
				if (this.session !== session) return;
				if (result === 'rejected') this.expire(SIGN_IN_STRINGS.ended);
				else this.scheduleCheck(config, PROBE_AGAIN_MS);
			});
		}, delay);
	}

	private armExpiry(session: TeamsSession): void {
		clearTimeout(this.expiryTimer);
		const wait = Math.min(Math.max(session.expiresAt - Date.now(), 0), MAX_TIMER_MS);
		this.expiryTimer = setTimeout(() => {
			if (this.session === session) this.expire(SIGN_IN_STRINGS.expired);
		}, wait);
	}

	private drop(): void {
		this.session = null;
		this.error = '';
		this.retryAfterSec = 0;
		clearTimeout(this.expiryTimer);
		clearTimeout(this.checkTimer);
		this.checkTimer = undefined;
		this.phase = 'required';
	}

	/** Forget everything (the element is leaving the page). */
	dispose(): void {
		this.generation++;
		this.probing = '';
		clearTimeout(this.expiryTimer);
		clearTimeout(this.checkTimer);
		this.checkTimer = undefined;
	}
}
