---
title: Compte et connexion
description: Ce que affiche par défaut Fichier > Compte, et comment brancher un flux de connexion sur son point d'accroche désactivé par défaut.
---

# Compte et connexion

L'ouverture et l'enregistrement de présentations hébergées par Microsoft est une intégration optionnelle de l'hôte. Voir
[Intégration OneDrive](/fr/guide/onedrive) pour l'adaptateur de stockage, les responsabilités d'authentification,
et [fichiers de test contre acceptation avec un vrai compte](/fr/guide/onedrive-testing).

Fichier > Compte contient un éditeur de profil local, un panneau stockage/confidentialité affichant des chiffres de stockage local en temps réel, une section À propos et une section de connexion facultative qui reste masquée tant qu'un hôte ne l'a pas explicitement branchée.

## Ce qui est fourni par défaut {#what-ships-by-default}

**Profil.** Un champ de nom d'affichage et une rangée de pastilles de couleur d'avatar. Rien n'est envoyé nulle part : c'est purement cosmétique, enregistré dans `localStorage` (clé `pptx-viewer-prefs`) avec le repli thème/langue de [Thèmes](/fr/guide/theming).

**Stockage et confidentialité.** Un compte en temps réel des présentations disposant d'un instantané local de sauvegarde automatique/récupération, et leur taille totale, lu depuis le même magasin IndexedDB que le flux de récupération au rechargement. Un bouton « Effacer les données locales » (protégé par une confirmation) supprime chaque instantané local et chaque préférence enregistrée.

**À propos.** Le nom du package du binding et sa version.

**Connexion.** Absente par défaut. S'affiche uniquement lorsque vous passez `accountAuth` avec `enabled: true` - voir ci-dessous.

Rien de tout cela ne nécessite de configuration. Le fonctionnement est identique dans une visionneuse nue, sans configuration, et dans une application hôte entièrement branchée.

## Brancher un vrai flux de connexion {#wiring-a-real-sign-in-flow}

La visionneuse n'a pas d'avis sur l'authentification : pas de client OAuth, pas de cookies de session, pas d'appel à un backend. `accountAuth` est un simple point d'accroche : lorsqu'il est activé, la page Compte affiche une invite de connexion et rappelle votre application ; c'est à vous de définir ce que signifie « connecté », puis de renvoyer le résultat vers le bas.

```ts
interface AccountAuthConfig {
	enabled: boolean;
	onSignIn: () => void;
	signedInUser?: {
		name: string;
		email?: string;
		avatarUrl?: string;
	};
}
```

- `enabled: false` (ou l'omission complète de `accountAuth`) est la valeur par défaut : la section de connexion ne s'affiche pas du tout.
- `onSignIn` se déclenche lorsque l'utilisateur clique sur le bouton de connexion. Lancez ici le flux réel de votre application (redirection, fenêtre OAuth, ouverture de votre propre modale, etc.) : la visionneuse n'attend pas et ne gère aucun état de chargement.
- Lorsque votre application sait qui est connecté, renvoyez `signedInUser` par la même prop `accountAuth` au rendu ou à la mise à jour suivante : la page Compte affiche alors « Connecté en tant que {name} » à la place du bouton de connexion.

```tsx
// React
const [user, setUser] = useState<{ name: string; email?: string } | undefined>();

<PowerPointViewer
	content={bytes}
	accountAuth={{
		enabled: true,
		onSignIn: () => startOAuthFlow().then(setUser),
		signedInUser: user,
	}}
/>;
```

```vue
<!-- Vue -->
<PowerPointViewer
	:content="bytes"
	:accountAuth="{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }"
/>
```

```html
<!-- Angular -->
<pptx-viewer
	[content]="bytes"
	[accountAuth]="{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }"
/>
```

```ts
// Vanilla
const viewer = createPptxViewer(host, {
	source: bytes,
	accountAuth: { enabled: true, onSignIn: startOAuthFlow, signedInUser: user },
});
```

```svelte
<!-- Svelte -->
<PowerPointViewer source={content} accountAuth={{ enabled: true, onSignIn: startOAuthFlow, signedInUser: user }} />
```

La forme est identique dans les cinq bindings.

## Étapes suivantes {#next-steps}

- [Thèmes](/fr/guide/theming) - le même repli `localStorage` alimente à la fois l'éditeur de profil et les sélecteurs Apparence/Langue.
- [Localisation (i18n)](/fr/guide/localization) - chaque chaîne de la page Compte passe par les clés de traduction standard `pptx.account.*`.
