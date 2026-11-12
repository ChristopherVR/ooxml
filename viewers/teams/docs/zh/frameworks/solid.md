# Solid

`openteams-solid-viewer`（对等依赖：`solid-js` 1.9）通过一个普通函数提供该元素，并以信号的形式提供原始客户端。

```bash
npm install openteams-solid-viewer solid-js
```

## 组件 {#the-component}

`Teams(props)` 返回 `<teams-app>` 元素，因此该包本身无需 JSX 转换；在你的 JSX 中直接使用即可：

```tsx
import type { TeamsServerConfig } from 'ooxml-core/teams';
import { Teams } from 'openteams-solid-viewer';

export function Workspace(props: { config: TeamsServerConfig }) {
	return (
		<div style={{ height: '100%' }}>
			{Teams({
				workspaceId: 'acme',
				userName: 'Ada',
				get config() {
					return props.config;
				},
				onOpenFile: (detail) => console.log(detail.url),
			})}
		</div>
	);
}
```

属性会以响应式方式重新应用（对于响应式的值请使用 getter），事件处理函数则在事件触发时读取。

## 原语 {#the-primitive}

```tsx
import { createTeamsClient } from 'openteams-solid-viewer';

const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user, config }));
// state()?.channels, client()?.send({ text: 'hi' })
```

当 `options()` 变化时客户端会被重新创建，并随其所属的拥有者一同销毁。
