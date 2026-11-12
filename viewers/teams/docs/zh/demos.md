# 在线演示

本站为每个绑定发布一个演示。所有演示都运行同一个工作空间，并接受相同的查询参数（`?name=Ada&room=acme`）。

| 演示                                                                                                 | 展示内容                                                      | 源码                  |
| ---------------------------------------------------------------------------------------------------- | ------------------------------------------------------------- | --------------------- |
| [原生 JavaScript](/demo/){target="_self"}（也位于 [/demo-vanilla/](/demo-vanilla/){target="_self"}） | 完整应用：由 `mountTeams` 挂载的一个 `<teams-app>`。          | `demos/teams/vanilla` |
| [React](/demo-react/){target="_self"}                                                                | 左侧为 `<Teams />`，右侧为基于 `useTeams()` hook 手写的面板。 | `demos/teams/react`   |
| [Vue](/demo-vue/){target="_self"}                                                                    | `openteams-vue-viewer` 的 `<Teams>` 组件。                    | `demos/teams/vue`     |
| [Angular](/demo-angular/){target="_self"}                                                            | `openteams-angular-viewer` 的 `<teams-workspace>`。           | `demos/teams/angular` |
| [Svelte](/demo-svelte/){target="_self"}                                                              | `openteams-svelte-viewer` 的 `Teams.svelte`。                 | `demos/teams/svelte`  |
| [Solid](/demo-solid/){target="_self"}                                                                | `openteams-solid-viewer` 的 `Teams()`。                       | `demos/teams/solid`   |

由于它们共享同一个源（origin）和同一个本地房间名，一个框架的演示可以与另一个框架的演示通讯：在两个标签页中分别打开 `/demo-vue/?name=Ada&room=x` 和 `/demo-svelte/?name=Bob&room=x` 即可。

## 无需服务器即可运行 {#they-run-without-a-server}

GitHub Pages 只提供静态文件，因此这些页面背后没有服务器，也不会伪装出一个服务器。演示以 `VITE_TEAMS_STATIC=1` 构建，这会让它们以核心的**本地模式**启动：

- 聊天、在线状态和通话信令只在**你的浏览器内部**的标签页之间传递，通过 `BroadcastChannel` 实现；
- 共享文档保存在你浏览器的 `localStorage` 中，因此刷新页面后依然保留；
- 文件仅按名称共享，因为没有地方可以上传它们；
- 通话只连接同一浏览器中的标签页（使用你的摄像头和麦克风，需经你授权）。

每个演示都会显示一条说明此情况的提示，并附带一个链接，可打开第二个标签页，以另一个人的身份加入（即 `?name=` 参数）。在一个标签页中输入，消息会出现在另一个标签页中。

如果你在原生 JavaScript 演示的**服务器**设置中保存了自己的服务器，演示将改用该服务器；在那里清除它（或清除站点存储）即可恢复本地模式。

## 使用你自己的服务器 {#use-your-own-server}

真正的使用需要服务器：运行 `npx openteams-server --insecure-dev`（仅用于本地开发），并参阅[自带服务器](/zh/server)。从源码运行时，`bun run dev` 会启动参考服务器，并让两个演示都连接到它（参见[快速开始](/zh/getting-started#run-it-from-source)）。
