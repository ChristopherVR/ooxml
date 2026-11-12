# React

`openteams-react-viewer`（对等依赖：`react` 18 或更高版本）以 `<Teams />` 的形式提供完整界面，并以 hook 的形式提供原始客户端。

```bash
npm install openteams-react-viewer react
```

## 组件 {#the-component}

```tsx
import { Teams } from 'openteams-react-viewer';

export function Workspace({ config }) {
	return (
		<Teams
			workspaceId="acme"
			userName="Ada"
			userId="u-ada"
			config={config}
			onOpenFile={(detail, event) => {
				if (detail.attachment.kind === 'docx') {
					event.preventDefault();
					openInWord(detail.url);
				}
			}}
		/>
	);
}
```

属性会在每次渲染后应用，且只有发生变化的值才会传递给元素。ref 即 `<teams-app>` 元素本身（其 `client` getter 返回核心客户端）。

## Hook {#the-hooks}

```tsx
import { useTeams, type TeamsClientOptions } from 'openteams-react-viewer';

function Panel({ options }: { options: TeamsClientOptions }) {
	const { client, state } = useTeams(options);
	if (!client || !state) return <p>Connecting…</p>;
	return (
		<ul>
			{state.channels.map((c) => (
				<li key={c.id}>
					<button onClick={() => client.select(c.id)}># {c.name}</button>
				</li>
			))}
		</ul>
	);
}
```

`useTeams` 等同于 `useTeamsClient` 加上 `useTeamsState`（基于 `useSyncExternalStore`）。当工作空间、用户或服务器发生变化时，客户端会被重新创建，并在卸载时销毁；请保持 `options` 稳定（使用 `useMemo`）。

[React 演示](/demo-react/){target="_self"}在同一页面上同时使用两者：`demos/teams/react/src/App.tsx`。
