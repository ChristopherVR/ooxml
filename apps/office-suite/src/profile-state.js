const key = 'ooxml-suite-profiles-v1';
const changed = new EventTarget();
const validId = (id) => typeof id === 'string' && /^[\w-]{1,100}$/u.test(id);
let state;
try {
	const saved = JSON.parse(localStorage.getItem(key));
	if (
		saved?.profiles?.length &&
		saved.profiles.every((p) => validId(p.id) && typeof p.name === 'string')
	)
		state = saved;
} catch {}
if (!state) {
	let id;
	try {
		id = localStorage.getItem('suite-user-id');
	} catch {}
	state = {
		active: 'default',
		profiles: [
			{
				id: 'default',
				userId: validId(id) ? id : crypto.randomUUID(),
				name: 'Your profile',
				email: '',
				organization: 'Personal workspace',
				title: '',
				photo: '',
			},
		],
	};
}
if (!state.profiles.some((p) => p.id === state.active)) state.active = state.profiles[0].id;
try {
	localStorage.setItem(key, JSON.stringify(state));
} catch {}
export const profiles = () => state.profiles.map((p) => ({ ...p }));
export const currentProfile = () => ({ ...state.profiles.find((p) => p.id === state.active) });
export const documentDatabase = () =>
	state.active === 'default' ? 'ooxml-suite-documents' : `ooxml-suite-documents-${state.active}`;
export const teamsWorkspace = () =>
	state.active === 'default' ? 'office-suite' : `suite-${state.active}`;
export const initials = (name) =>
	name
		.trim()
		.split(/\s+/u)
		.slice(0, 2)
		.map((s) => s[0])
		.join('')
		.toUpperCase();
export const profilePhoto = (p) =>
	typeof p.photo === 'string' &&
	/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/u.test(p.photo)
		? p.photo
		: '';
function persist(next) {
	localStorage.setItem(key, JSON.stringify(next));
	state = next;
	changed.dispatchEvent(new Event('change'));
}
export function updateProfile(values) {
	const name = String(values.name ?? '')
		.trim()
		.slice(0, 80);
	if (!name) throw new Error('Enter a display name.');
	const p = {
		...currentProfile(),
		name,
		email: String(values.email ?? '')
			.trim()
			.slice(0, 160),
		organization:
			String(values.organization ?? '')
				.trim()
				.slice(0, 100) || 'Personal workspace',
		title: String(values.title ?? '')
			.trim()
			.slice(0, 100),
		photo: profilePhoto(values),
	};
	persist({
		...state,
		profiles: state.profiles.map((existing) => (existing.id === p.id ? p : existing)),
	});
}
export function addProfile(name) {
	name = name.trim().slice(0, 80);
	if (!name) throw new Error('Enter a profile name.');
	const p = {
		id: crypto.randomUUID(),
		userId: crypto.randomUUID(),
		name,
		organization: 'Personal workspace',
		email: '',
		title: '',
		photo: '',
	};
	persist({ ...state, profiles: [...state.profiles, p] });
	return p;
}
export function selectProfile(id) {
	if (!state.profiles.some((p) => p.id === id)) throw new Error('Profile not found.');
	persist({ ...state, active: id });
}
export function onProfile(listener) {
	const handle = () => listener(currentProfile());
	changed.addEventListener('change', handle);
	handle();
	return () => changed.removeEventListener('change', handle);
}
