import { $, escape, choose } from './ui.js';
import { icon } from './icons.js';
import {
	currentProfile,
	profiles,
	initials,
	profilePhoto,
	updateProfile,
	addProfile,
	selectProfile,
	onProfile,
} from './profile-state.js';

export function avatar(p, large = false) {
	return `<span class="avatar${large ? ' avatar-large' : ''}">${profilePhoto(p) ? `<img src="${profilePhoto(p)}" alt="" />` : escape(initials(p.name))}</span>`;
}

export function mountProfile({ beforeSwitch, changed }) {
	const menu = $('account-menu');
	const workspace = $('workspace-menu');
	let switching = false;
	async function switchTo(id) {
		if (id === currentProfile().id || switching) return;
		switching = true;
		try {
			await beforeSwitch();
		} catch (error) {
			switching = false;
			throw error;
		}
		selectProfile(id);
		location.replace(`${location.pathname}${location.search}#/`);
		location.reload();
	}
	function render(p) {
		$('profile-toggle').innerHTML = avatar(p);
		$('profile-toggle').setAttribute('aria-label', `Account manager for ${p.name}`);
		$('workspace-name').textContent = p.organization;
		$('welcome').textContent =
			p.name === 'Your profile'
				? 'A place for everything you are working on.'
				: `Welcome back, ${p.name.split(' ')[0]}.`;
		$('workspace-status').textContent = `${p.organization} · This device`;
		menu.innerHTML = `<div class="account-topline"><span>OOXML Office</span><span class="local-label">Local profile</span></div>
		<div class="account-identity">${avatar(p, true)}<div><h2>${escape(p.name)}</h2><p>${escape(p.email || 'No connected account')}</p>${p.title ? `<p>${escape(p.title)}</p>` : ''}<button class="text-link" data-profile-edit>View my profile</button></div></div>
		<div class="account-workspace">${icon('folder')}<div><strong>${escape(p.organization)}</strong><small>Files stored on this device</small></div></div>
		<div class="account-actions"><button data-workspace-switch>${icon('folder')}Switch workspace</button><button data-appearance>${icon('settings')}Appearance & themes</button></div>
		<p class="account-feedback" role="alert" hidden></p><p class="account-footnote">Local profiles keep separate libraries in this browser. They are not secured sign-in accounts.</p>`;
		workspace.innerHTML = `<div class="workspace-menu-heading"><h2>Switch workspace</h2><p>Each local profile has its own library and Teams workspace.</p></div><div class="profile-switcher">${profiles()
			.map(
				(other) =>
					`<button data-switch-profile="${other.id}" ${other.id === p.id ? 'aria-current="true"' : ''}>${avatar(other)}<span><strong>${escape(other.organization)}</strong><small>${escape(other.name)}</small></span><span>${other.id === p.id ? 'Active' : 'Switch'}</span></button>`,
			)
			.join(
				'',
			)}<button data-profile-add>${icon('add')}Add a local workspace</button></div><p class="account-feedback" role="alert" hidden></p>`;
		changed?.(p);
	}
	onProfile(render);
	menu.addEventListener('toggle', () =>
		$('profile-toggle').setAttribute('aria-expanded', String(menu.matches(':popover-open'))),
	);
	document.addEventListener('click', (event) => {
		const b = event.target.closest('button');
		if (!b) return;
		if (b.hasAttribute('data-workspace-switch')) {
			menu.hidePopover();
			workspace.showPopover();
		}
		if (b.hasAttribute('data-profile-edit')) {
			menu.hidePopover();
			edit();
		}
		if (b.hasAttribute('data-profile-add')) {
			workspace.hidePopover();
			add();
		}
		if (b.hasAttribute('data-appearance')) {
			menu.hidePopover();
			$('theme-picker').click();
		}
		if (b.dataset.switchProfile)
			void switchTo(b.dataset.switchProfile).catch((error) => {
				const message = workspace.querySelector('.account-feedback');
				message.hidden = false;
				message.textContent = error.message;
			});
	});
	function edit() {
		const p = currentProfile();
		let photo = profilePhoto(p);
		const form = document.createElement('form');
		form.className = 'profile-form';
		form.innerHTML = `<div class="profile-form-identity"><div id="photo-preview">${avatar(p, true)}</div><div><label class="photo-label" for="profile-photo">Change photo</label><input id="profile-photo" class="sr-only" name="photo" type="file" accept="image/png,image/jpeg,image/webp" /><p class="photo-help">JPG, PNG or WebP, up to 8 MB</p><button type="button" class="text-link" data-remove-photo>Remove photo</button></div></div>
		<div class="profile-fields"><label>Display name<input name="name" value="${escape(p.name)}" maxlength="80" required autocomplete="name" /></label><label>Email address<input name="email" type="email" value="${escape(p.email)}" maxlength="160" autocomplete="email" /><small>Contact information, not a sign-in.</small></label><label>Workspace name<input name="organization" value="${escape(p.organization)}" maxlength="100" required /><small>The name shown in your sidebar.</small></label><label>Job title (optional)<input name="title" value="${escape(p.title)}" maxlength="100" /></label></div>
		<p class="profile-privacy">Your display name is also used in Teams. Your email and photo stay in this browser.</p><p class="form-error" role="alert"></p><div class="form-footer"><button type="button" data-cancel>Cancel</button><button class="primary">Save profile</button></div>`;
		let photoLoading = false;
		const submit = form.querySelector('.form-footer .primary');
		const failure = (error) => {
			form.querySelector('.form-error').textContent = error.message;
		};
		form.querySelector('[name=photo]').onchange = async (event) => {
			try {
				const file = event.target.files[0];
				if (!file) return;
				photoLoading = true;
				submit.disabled = true;
				if (
					!['image/png', 'image/jpeg', 'image/webp'].includes(file.type) ||
					file.size > 8 * 1024 * 1024
				)
					throw new Error('Choose a PNG, JPEG or WebP photo smaller than 8 MB.');
				const candidate = await new Promise((resolve, reject) => {
					const reader = new FileReader();
					reader.onload = () => resolve(reader.result);
					reader.onerror = reject;
					reader.readAsDataURL(file);
				});
				const image = new Image();
				image.src = candidate;
				await image.decode();
				const canvas = document.createElement('canvas');
				const ratio = Math.min(1, 256 / Math.max(image.naturalWidth, image.naturalHeight));
				canvas.width = Math.max(1, Math.round(image.naturalWidth * ratio));
				canvas.height = Math.max(1, Math.round(image.naturalHeight * ratio));
				canvas.getContext('2d').drawImage(image, 0, 0, canvas.width, canvas.height);
				photo = canvas.toDataURL('image/png');
				form.querySelector('#photo-preview').innerHTML = avatar({ ...p, photo }, true);
				form.querySelector('.form-error').textContent = '';
			} catch (error) {
				failure(error);
			} finally {
				photoLoading = false;
				submit.disabled = false;
				event.target.value = '';
			}
		};
		form.querySelector('[data-remove-photo]').onclick = () => {
			if (photoLoading) return;
			photo = '';
			form.querySelector('#photo-preview').innerHTML = avatar({ ...p, photo }, true);
		};
		form.querySelector('[data-cancel]').onclick = () => $('dialog').close();
		form.onsubmit = (event) => {
			event.preventDefault();
			try {
				if (photoLoading) return;
				updateProfile({ ...Object.fromEntries(new FormData(form)), photo });
				$('dialog').close();
			} catch (error) {
				failure(error);
			}
		};
		choose('My profile', form);
	}
	function add() {
		const form = document.createElement('form');
		form.innerHTML =
			'<p>Create a separate local library and Teams workspace. Your current files remain in your existing profile.</p><label>Profile name<input name="name" required maxlength="80" placeholder="e.g. Personal or Design studio" /></label><p class="form-error" role="alert"></p><button class="primary">Create profile</button>';
		form.onsubmit = async (event) => {
			event.preventDefault();
			const submit = form.querySelector('button');
			submit.disabled = true;
			try {
				await beforeSwitch();
				const p = addProfile(new FormData(form).get('name'));
				selectProfile(p.id);
				location.replace(`${location.pathname}${location.search}#/`);
				location.reload();
			} catch (error) {
				form.querySelector('.form-error').textContent = error.message;
				submit.disabled = false;
			}
		};
		choose('Add a local profile', form);
	}
}
