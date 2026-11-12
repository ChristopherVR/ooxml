---
title: Cuenta e inicio de sesión
description: Lo que muestra Archivo > Cuenta de forma predeterminada, y cómo conectar un flujo de inicio de sesión a su punto de enlace, desactivado de forma predeterminada.
---

# Cuenta e inicio de sesión

Abrir y guardar presentaciones alojadas en Microsoft es una integración opcional con el host. Consulte [Integración con OneDrive](/es/guide/onedrive) para el adaptador de almacenamiento y las responsabilidades de autenticación, y [pruebas con datos de prueba frente a aceptación con una cuenta real](/es/guide/onedrive-testing).

Archivo > Cuenta contiene un editor de perfil local, un panel de almacenamiento y privacidad que muestra cifras en vivo del almacenamiento local, una sección Acerca de y una sección de inicio de sesión opcional que permanece oculta hasta que un host la configura explícitamente.

## Lo que incluye de forma predeterminada {#what-ships-by-default}

**Perfil.** Un campo para el nombre y una fila de muestras de color para el avatar. No se envía nada a ningún sitio: es puramente cosmético y se guarda en `localStorage` (clave `pptx-viewer-prefs`) junto con el respaldo de tema e idioma de [Temas](/es/guide/theming).

**Almacenamiento y privacidad.** Un recuento en vivo de las presentaciones con una copia local de autoguardado y recuperación, y su tamaño total, leídos del mismo almacén de IndexedDB que usa el flujo de recuperación al recargar. El botón "Borrar datos locales" (con confirmación) elimina todas las copias locales y todas las preferencias guardadas.

**Acerca de.** El nombre del paquete y la versión del enlace.

**Inicio de sesión.** No aparece de forma predeterminada. Solo se muestra cuando pasa `accountAuth` con `enabled: true`; consulte más abajo.

Nada de esto requiere configuración. Funciona igual en un visualizador sin configurar que en una aplicación host completamente integrada.

## Cómo conectar un flujo de inicio de sesión real {#wiring-a-real-sign-in-flow}

El visualizador no decide nada sobre la autenticación: no incluye cliente OAuth, cookies de sesión ni llamadas a un backend. `accountAuth` es un gancho sencillo: cuando está activado, la página de Cuenta muestra una solicitud de inicio de sesión y llama a su aplicación. Usted decide qué significa "sesión iniciada" y devuelve el resultado al visualizador.

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

- `enabled: false` (o omitir `accountAuth` por completo) es el valor predeterminado: la sección de inicio de sesión no se renderiza.
- `onSignIn` se dispara al hacer clic en el botón de inicio de sesión. Inicie aquí el flujo real de su aplicación (redirección, ventana emergente de OAuth, abrir su propio modal, etc.). El visualizador no espera a que termine ni gestiona ningún estado de carga.
- Cuando su aplicación sepa quién ha iniciado sesión, pase `signedInUser` de vuelta a través de la misma prop `accountAuth` en la siguiente renderización o actualización. La página de Cuenta mostrará entonces "Sesión iniciada como {name}" en lugar del botón de inicio de sesión.

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

La forma es idéntica en los cinco enlaces.

## Próximos pasos {#next-steps}

- [Temas](/es/guide/theming): el mismo respaldo de `localStorage` sirve al editor de perfil y a los selectores de Apariencia e Idioma.
- [Localización (i18n)](/es/guide/localization): cada cadena de la página de Cuenta pasa por las claves de traducción estándar `pptx.account.*`.
