# Las demos en vivo

Con este sitio se publica una demo por binding. Todas ejecutan el mismo espacio de trabajo y aceptan
los mismos parámetros de consulta (`?name=Ada&room=acme`).

| Demo                                                                                            | Qué muestra                                                                                   | Código fuente         |
| ----------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- | --------------------- |
| [Vanilla](/demo/){target="_self"} (también en [/demo-vanilla/](/demo-vanilla/){target="_self"}) | La aplicación completa: un único `<teams-app>` montado con `mountTeams`.                      | `demos/teams/vanilla` |
| [React](/demo-react/){target="_self"}                                                           | `<Teams />` a la izquierda y un panel escrito a mano sobre el hook `useTeams()` a la derecha. | `demos/teams/react`   |
| [Vue](/demo-vue/){target="_self"}                                                               | El componente `<Teams>` de `openteams-vue-viewer`.                                            | `demos/teams/vue`     |
| [Angular](/demo-angular/){target="_self"}                                                       | `<teams-workspace>` de `openteams-angular-viewer`.                                            | `demos/teams/angular` |
| [Svelte](/demo-svelte/){target="_self"}                                                         | `Teams.svelte` de `openteams-svelte-viewer`.                                                  | `demos/teams/svelte`  |
| [Solid](/demo-solid/){target="_self"}                                                           | `Teams()` de `openteams-solid-viewer`.                                                        | `demos/teams/solid`   |

Como comparten un mismo origen y un mismo nombre de sala local, una demo de un framework puede
hablar con una demo de otro: abre `/demo-vue/?name=Ada&room=x` y `/demo-svelte/?name=Bob&room=x` en
dos pestañas.

## Funcionan sin servidor {#they-run-without-a-server}

GitHub Pages solo sirve archivos estáticos, así que no hay ningún servidor detrás de estas páginas y
nada se simula para parecerlo. Las demos se compilan con `VITE_TEAMS_STATIC=1`, lo que las inicia en
el **modo local** del núcleo:

- el chat, la presencia y la señalización de llamadas viajan entre pestañas de **tu propio
  navegador**, mediante `BroadcastChannel`;
- el documento compartido se guarda en el `localStorage` de tu navegador, así que sobrevive a una
  recarga;
- los archivos se comparten solo por nombre, porque no hay dónde subirlos;
- las llamadas conectan pestañas del mismo navegador (tu cámara y tu micrófono, con tu permiso).

Cada demo muestra un aviso que lo indica, con un enlace que abre una segunda pestaña como otra
persona (el parámetro `?name=`). Escribe en una pestaña y el mensaje aparece en la otra.

Si guardaste tu propio servidor en los ajustes **Server** de la demo vanilla, la demo usa ese
servidor; bórralo allí (o borra el almacenamiento del sitio) para volver al modo local.

## Usa tu propio servidor {#use-your-own-server}

El uso real necesita un servidor: ejecuta `npx openteams-server --insecure-dev` (solo para desarrollo
local) y consulta [Aporta tu propio servidor](/es/server). Desde el código fuente, `bun run dev` inicia
el servidor de referencia y ambas demos conectadas a él (consulta
[Primeros pasos](/es/getting-started#run-it-from-source)).
