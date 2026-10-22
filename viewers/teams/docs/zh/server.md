# 自带服务器

OpenTeams 没有托管服务。工作空间要么**不使用服务器**运行（`mode: 'local'`，仅限同一浏览器的标签页），要么连接到**你自己运行的服务器**，该服务器需遵守三个小型约定。参考实现 `openteams-server` 在一个约 300 行的 Node 进程中完成这三项；任何遵守相同约定的其他服务器同样可用。

::: danger 无端到端加密
聊天以 Yjs 更新的形式经由你的服务器传输，服务器会以明文形式存储房间文档和每一个上传的文件，并且可以读取它们。通话媒体在对等端之间加密（DTLS-SRTP），TURN 中继无法解密，但信令会经过服务器。
:::

## 运行参考服务器

需要 Node 22 或更高版本。

```bash
npx openteams-server                    # listens on 127.0.0.1:8787
TEAMS_TOKEN=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
```

它仅通过环境变量配置；该命令不接受任何命令行参数。

| 变量            | 默认值      | 含义                                                                                                    |
| --------------- | ----------- | ------------------------------------------------------------------------------------------------------- |
| `PORT`          | `8787`      | 监听端口。                                                                                              |
| `HOST`          | `127.0.0.1` | 监听地址。仅在 TLS 之后才绑定 `0.0.0.0`。                                                               |
| `TEAMS_TOKEN`   | 无          | 若设置，套接字通过 `?token=` 携带它，文件请求通过 `Authorization: Bearer` 携带它。                      |
| `TEAMS_ORIGINS` | 无          | 以逗号分隔的、允许访问套接字和文件的 `Origin`。未设置时：有令牌则允许任意来源，无令牌则仅允许回环地址。 |
| `TEAMS_DATA`    | `./data`    | 数据目录：`rooms/<room>.bin` 快照和 `files/<workspace>/`。                                              |

服务器还会读取用于账户的变量（`TEAMS_USERS_FILE`、`TEAMS_SESSION_SECRET`、`TEAMS_ADMIN_TOKEN`、`TEAMS_AUTH_RATE_*`）和用于限制的变量（`TEAMS_INSECURE`、`TEAMS_MAX_CONNECTIONS`、`TEAMS_MAX_CONNECTIONS_PER_IP`、`TEAMS_MAX_ROOMS`、`TEAMS_MAX_DOC_MB`、`TEAMS_MESSAGE_RATE`、`TEAMS_MESSAGE_BURST`、`TEAMS_FILES_QUOTA_MB`）。限制见[部署指南](/zh/deploy)；服务器的 README 是完整参考。对应的 `createTeamsServer` 选项有 `users`、`authenticate`、`sessionSecret`、`adminToken`、`sessionRate`、`authorizeSync`、`maxConnections`、`maxRooms`、`messageRate`、`filesQuotaBytes` 和 `insecure`。

也可以将其嵌入：

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({
	token: process.env.TEAMS_TOKEN,
	origins: ['https://app.example.com'],
	dataDir: '/var/lib/openteams',
	maxFileBytes: 25 * 1024 * 1024, // default
	log: console.log,
});
await server.listen(8787, '127.0.0.1');
// server.http is the node:http server; server.close() flushes every room to disk.
```

然后让客户端指向它：

```ts
const config = {
	mode: 'server',
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	token: 'change-me', // only if TEAMS_TOKEN is set
};
```

[部署参考服务器](/zh/deploy)介绍了 TLS、令牌、TURN、Docker 以及一个 Caddy 示例。

## 端点

| 端点                                  | 作用                                                                                                      |
| ------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `ws /sync/<room>`                     | 兼容 y-websocket 的同步：每个房间一个 `Y.Doc` 和 awareness，持久化到 `rooms/<room>.bin`。消息最大 32 MB。 |
| `ws /signal/<room>`                   | 通话信令的 JSON 中继：将每条文本消息转发给房间内的其他套接字。消息最大 256 KB。                           |
| `POST /files/<workspace>/<name>`      | 存储附件（上限 25 MB），并返回 `201 { url, size }`。                                                      |
| `GET /files/<workspace>/<name>`       | 以附件形式下载（`Content-Disposition: attachment`、`Referrer-Policy: no-referrer`、`nosniff`）。          |
| `POST /files/link/<workspace>/<name>` | 返回 `{ url }`：仅针对该文件签名的链接，有效期五分钟，因此令牌不会出现在 URL 中。                         |
| `GET /health`                         | `{ ok, rooms, calls, auth }`。                                                                            |
| `POST /auth/session`                  | 为 `{ user, credential }` 签发会话（有速率限制）。参见服务器 README 中的 "Accounts and sessions"。        |
| `POST /auth/check`, `/auth/signout`   | 检查 `{ session }`（204 或 401）；登出，会撤销该用户的会话。                                              |
| `POST /auth/revoke`                   | 运营者使用 `TEAMS_ADMIN_TOKEN` 撤销 `{ user }`；没有管理员令牌时为 404。                                  |

房间名必须匹配 `^[\w-]{1,128}$`，文件路径中的工作空间 id 必须匹配 `^[\w-]{1,64}$`；文件名会被清理为仅含字母、数字、`.`、`-`、`_`、空格和括号，并截断为 120 个字符。

## 约定

任何完成这三件事的服务器都可以配合所有绑定使用。

### 同步

`ws(s)://host/<base>/<room>` 使用 y-websocket 的线路格式（同步步骤 1 和 2、更新、awareness、查询 awareness）。标准的 y-websocket 服务器即可使用。为每个房间维护一个 `Y.Doc` 并将其持久化的服务器，可以让之后加入的人看到历史记录。

### 信令

`ws(s)://host/<base>/<room>` 将每条 JSON 文本消息转发给房间内的**其他**套接字，并在某个套接字关闭时向它们发送 `{ "type": "bye", "from": <id> }`。一个套接字对应一个对等端 id（它使用的第一个 `from`）。服务器无需理解这些消息。

### 文件 {#files}

可选。使用 `POST|GET <origin>/files/<workspace>/<name>`，令牌放在 `Authorization: Bearer` 头中，另有 `POST <origin>/files/link/<workspace>/<name>` 返回短期有效的签名 `{ url }`，用于在浏览器标签页中打开文件。origin 取自 `syncUrl`。没有文件服务器时，可以给元素提供一个 `uploadFile` 函数（把附件放入你自己的对象存储），否则附件仅按名称共享。

配置了令牌时，两个套接字都会收到 `?token=<token>`，因为浏览器无法为 WebSocket 设置请求头。请在服务器端强制校验它以及来源白名单。

## 通话：STUN 与 TURN {#calls-stun-and-turn}

`iceServers` 承载 STUN 和 TURN 的 URL。公共 STUN 在开放网络上足够使用，但在对称 NAT 或严格防火墙之后永远不够：请运行你自己的 TURN（例如 coturn），并向客户端提供凭据。没有凭据的 TURN 条目会被核心的 `parseServerConfig` 拒绝。静态 TURN 凭据对任何能打开应用的人都是可见的；对于任何公开场景，请签发短期凭据。[部署指南](/zh/deploy#calls-across-nats-turn)中有一个 coturn 示例。

## 参考服务器不做的事情 {#what-the-reference-server-does-not-do}

在让人们使用它之前，请先弄清楚以下几点：

- **无端到端加密。** 它存储并且可以读取每一条消息和每一个文件。
- **账户是可选且精简的。** 使用 `users`（或 `TEAMS_USERS_FILE`）或 `authenticate` 函数时，服务器会签发会话（`POST /auth/session`）、从会话中获取用户 ID，并可以撤销会话。没有它们时，`TEAMS_TOKEN` 是所有人共用的一个密钥，保存在浏览器中，用户 ID 只是声明。没有密码重置、账户管理、SSO，也没有按团队的访问控制。客户端中的作者和角色检查可以被恶意客户端绕过，因此服务器是信任边界：请使用 `authenticate` 接入你的身份提供方，并用 `authorizeSync` 实现按房间的权限。
- **自身不提供 TLS。** 请在终止 TLS 并升级 WebSocket 的反向代理之后运行它；浏览器会拒绝 `https://` 页面发起的 `ws://` 连接。
- **单进程，房间驻留内存。** 它不支持水平扩展。房间在其最后一个套接字关闭 30 秒后被卸载。
- **每个房间一个快照，不做压缩。** 文档会随历史增长，且从不修剪。
- **简单的文件存储。** 没有列表、没有删除、没有按工作空间的配额；上传同名文件会替换原文件。每个文件上限 25 MB。
- **无速率限制**，也没有媒体服务器：通话是点对点的（网状）。
