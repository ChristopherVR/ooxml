import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import {
	DEFAULT_OFFICE_PROFILE,
	OFFICE_AVATAR_SWATCHES,
	profileInitials,
	readOfficeProfile,
	sanitizeOfficeProfile,
	writeOfficeProfile,
	type OfficeProfile,
} from '../controls';
import { registerOfficeUi } from '../index';

beforeAll(() => registerOfficeUi());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

type Account = HTMLElement & { profile: OfficeProfile };

describe('office profile', () => {
	it('derives initials and sanitizes untrusted data', () => {
		expect(profileInitials({ displayName: 'Ada Lovelace', avatarColor: '#000000' })).toBe('AL');
		expect(profileInitials(DEFAULT_OFFICE_PROFILE)).toBe('?');
		expect(sanitizeOfficeProfile({ displayName: 3, avatarColor: 'red;x', initial: '' })).toEqual(
			DEFAULT_OFFICE_PROFILE,
		);
	});

	it('round-trips through one suite-wide storage key and tolerates corrupt data', () => {
		writeOfficeProfile({ displayName: ' Ada ', avatarColor: '#16a34a' });
		expect(readOfficeProfile()).toEqual({ displayName: 'Ada', avatarColor: '#16a34a' });
		localStorage.setItem('ooxml-office-profile', '{oops');
		expect(readOfficeProfile()).toEqual(DEFAULT_OFFICE_PROFILE);
		const failing = {
			getItem: () => {
				throw new Error('denied');
			},
			setItem: () => {
				throw new Error('denied');
			},
			removeItem: () => {},
		};
		expect(readOfficeProfile(failing)).toEqual(DEFAULT_OFFICE_PROFILE);
		expect(() => writeOfficeProfile(DEFAULT_OFFICE_PROFILE, failing)).not.toThrow();
	});
});

describe('office-ui-account', () => {
	function mount() {
		const account = document.createElement('office-ui-account') as Account;
		document.body.append(account);
		account.profile = { displayName: 'Ada Lovelace', avatarColor: '#c2431f' };
		const root = account.shadowRoot!;
		const change = vi.fn();
		account.addEventListener('office-profile-change', (event) =>
			change((event as CustomEvent).detail.profile),
		);
		return { account, root, change };
	}

	it('shows the avatar, name and checked colour', () => {
		const { root } = mount();
		expect(root.querySelector('h2')!.textContent).toBe('User Information');
		expect(root.querySelector('.avatar')!.textContent).toBe('AL');
		expect(root.querySelector('.name')!.textContent).toBe('Ada Lovelace');
		expect(root.querySelector('input')!.value).toBe('Ada Lovelace');
		expect(root.querySelectorAll('[role="radio"]')).toHaveLength(OFFICE_AVATAR_SWATCHES.length);
		expect(root.querySelector('[aria-checked="true"]')!.getAttribute('data-color')).toBe('#c2431f');
	});

	it('emits profile changes from the name field, swatches and arrow keys', () => {
		const { root, change } = mount();
		const input = root.querySelector('input')!;
		input.value = 'Grace Hopper';
		input.dispatchEvent(new Event('change'));
		expect(change).toHaveBeenLastCalledWith({
			displayName: 'Grace Hopper',
			avatarColor: '#c2431f',
		});
		expect(root.querySelector('.avatar')!.textContent).toBe('GH');
		root.querySelector<HTMLButtonElement>('[data-color="#16a34a"]')!.click();
		expect(change).toHaveBeenLastCalledWith({
			displayName: 'Grace Hopper',
			avatarColor: '#16a34a',
		});
		root
			.querySelector('[role="radiogroup"]')!
			.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(change).toHaveBeenLastCalledWith({
			displayName: 'Grace Hopper',
			avatarColor: '#ca8a04',
		});
		expect(root.activeElement?.getAttribute('data-color')).toBe('#ca8a04');
	});
});
