# 运行你自己的服务器

## 参考服务器

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| 变量                                                    | 含义                                                                                                                                    |
| ------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`、`HOST`                                          | 监听地址。仅在 TLS 之后才绑定 `0.0.0.0`。                                                                                               |
| `TEAMS_TOKEN`                                           | 若设置，套接字通过 `?token=` 携带它，文件请求通过 `Authorization: Bearer` 携带它。                                                      |
| `TEAMS_ORIGINS`                                         | 以逗号分隔的、允许访问套接字和文件的 `Origin`。为空时：设置了令牌则允许任意来源；未设置令牌则仅允许回环来源和不发送 `Origin` 的客户端。 |
| `TEAMS_DATA`                                            | 数据目录（默认 `./data`）：`rooms/<room>.bin` 快照和 `files/`。                                                                         |
| `TEAMS_INSECURE`                                        | `1`（或 `--insecure`）允许在未设置令牌时使用非回环的 `HOST`。否则服务器拒绝启动。                                                       |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | 打开的套接字总数（默认 500）和每个远程地址的数量（20）。                                                                                |
| `TEAMS_MAX_ROOMS`                                       | 同步房间数和通话房间数（默认各 1000）。                                                                                                 |
| `TEAMS_MAX_DOC_MB`                                      | 持久化房间文档的最大体积，单位 MB（默认 64）；超出的更新会被拒绝。                                                                      |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | 每个套接字每秒的消息数（默认 100）和突发量（200）；超出的套接字会被关闭。                                                               |
| `TEAMS_FILES_QUOTA_MB`                                  | 文件目录的总大小，单位 MB（默认 1024）；超出的上传返回 507。                                                                            |

它提供 `ws /sync/<room>`（兼容 y-websocket，每个房间一个持久化的 `Y.Doc`）、`ws /signal/<room>`（JSON 中继）和 `POST|GET /files/<workspace>/<name>`（上限 25 MB，文件名经过清理，下载以附件形式发送并带有 `Referrer-Policy: no-referrer`）。配置了令牌时，浏览器通过 `POST /files/link/<workspace>/<name>`（令牌放在请求头中）打开文件，该请求返回一个仅针对该文件签名、有效期五分钟的链接。`GET /health` 报告房间数和通话数。

**请将它置于 TLS 之后。** 浏览器会拒绝 `https://` 页面发起的 `ws://` 连接，而通过明文 HTTP 传输的令牌或通话信令在网络上是可读的。任何能升级 WebSocket 的反向代理都可以（nginx、Caddy、Traefik）。Caddy 示例：

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

然后让应用指向 `wss://teams.example.com/sync` 和 `wss://teams.example.com/signal`。

**令牌是一个共享密钥**，保存在浏览器中：它可以把陌生人挡在小团队的服务器之外，但并不是按用户的身份验证。若需要真正的账户，请为服务器提供 `users`（或 `TEAMS_USERS_FILE`）或面向你的身份提供方的 `authenticate` 函数，并设置 `TEAMS_SESSION_SECRET`，使会话在重启后仍然有效（参见服务器 README 中的 "Accounts and sessions"）。三个约定保持不变。

## 安全模型

- **一个共享令牌可打开所有房间和所有文件。** 只有启用账户时才有按用户的身份，只有传入 `authorizeSync` 时才有按房间的权限。请把令牌当作团队密码对待。
- **房间内的所有成员彼此信任。** 服务器只转发 Yjs 字节，因此任何成员都可以编辑或删除他人的消息，并伪造作者
  ID 和角色。不要邀请你不愿把整个房间托付的人。
- **令牌通过 WebSocket 查询字符串传输**（`?token=`），代理和访问日志可能记录它，浏览器也可能把它保留在历史记录中。
  请把服务器置于 TLS 之后，并且不要记录查询字符串。文件从不在 URL 中携带令牌：它们使用 `Authorization`
  请求头或签名链接。
- **默认配置仅限本机。** 未设置令牌时，服务器会拒绝非回环的 `HOST`（除非设置 `TEAMS_INSECURE=1`），并且只接受回环来源，
  因此你访问的网站无法操控你本机上的服务器。任何共享部署都应设置 `TEAMS_TOKEN` 和 `TEAMS_ORIGINS`。
- **使用短期有效的 TURN 凭据**（见下文）；静态凭据对每个打开应用的人都可见。
- **限制**（连接数、房间数、文档大小、消息速率和文件存储，即上面的变量）可以约束滥用，但不能取代反向代理自身的限流。
- 套接字只能代表它自己声明的在线状态 ID 和通话对等端 ID 发言，定向的通话信令只会投递给该对等端。上传绝不会覆盖已有文件：
  已被占用的名称会改用新名称保存，并在响应的 `url` 中返回。

## 跨 NAT 通话：TURN {#calls-across-nats-turn}

通话是点对点的。公共 STUN 在开放网络上可用；在对称 NAT 或严格防火墙之后，你需要 TURN 中继，它同样只传输密文媒体（DTLS-SRTP）。一个最小的 coturn 配置（`turnserver.conf`）：

```
listening-port=3478
tls-listening-port=5349
realm=turn.example.com
fingerprint
lt-cred-mech
user=teams:CHANGE_ME
no-multicast-peers
cert=/etc/letsencrypt/live/turn.example.com/fullchain.pem
pkey=/etc/letsencrypt/live/turn.example.com/privkey.pem
```

然后在应用的服务器设置中，每行填写一个 ICE 服务器：

```
stun:turn.example.com:3478
turn:turn.example.com:3478,turns:turn.example.com:5349 teams CHANGE_ME
```

静态 TURN 凭据对任何能打开应用的人都是可见的。对于任何公开场景，请签发短期凭据（coturn 的 `use-auth-secret` 加一个签发它们的小型端点），并通过你提供给元素的 `iceServers` 传入。

## Docker

```dockerfile
FROM node:22-alpine
WORKDIR /app
COPY server/package.json ./
RUN npm install --omit=dev
COPY server/bin.mjs server/index.mjs server/auth.mjs server/passwords.mjs server/guards.mjs \
  server/rate-limit.mjs server/chat-access.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data
VOLUME /data
EXPOSE 8787
CMD ["node", "bin.mjs"]
```

请在 `viewers/teams` 中构建。它会复制入口文件导入的每个模块（即 `server/package.json` 的 `files`）；只复制 `index.mjs` 会因缺少模块而失败。该配方是通过在 Node 上用 `npm install --omit=dev` 运行这些文件来验证的，镜像本身未构建。

运行时请设置 `TEAMS_TOKEN` 和 `TEAMS_ORIGINS`，并置于你的 TLS 代理之后。

## 使用其他服务器

任何遵守核心 `docs/teams-area.md` 中约定的服务器都可以：同步使用标准的 `y-websocket` 服务器，信令使用任意 WebSocket 扇出，文件使用任意对象存储（给 `<teams-app>` 提供一个 `uploadFile` 函数，或在元素上设置 `uploadFile`，附件就会存放到那里）。
