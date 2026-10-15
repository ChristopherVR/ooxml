# 运行你自己的服务器

## 参考服务器

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| 变量            | 含义                                                                               |
| --------------- | ---------------------------------------------------------------------------------- |
| `PORT`、`HOST`  | 监听地址。仅在 TLS 之后才绑定 `0.0.0.0`。                                          |
| `TEAMS_TOKEN`   | 若设置，套接字通过 `?token=` 携带它，文件请求通过 `Authorization: Bearer` 携带它。 |
| `TEAMS_ORIGINS` | 以逗号分隔的、允许访问套接字和文件的 `Origin`。为空则允许任意来源。                |
| `TEAMS_DATA`    | 数据目录（默认 `./data`）：`rooms/<room>.bin` 快照和 `files/`。                    |

它提供 `ws /sync/<room>`（兼容 y-websocket，每个房间一个持久化的 `Y.Doc`）、`ws /signal/<room>`（JSON 中继）和 `POST|GET /files/<workspace>/<name>`（上限 25 MB，文件名经过清理，下载以附件形式发送并带有 `Referrer-Policy: no-referrer`）。配置了令牌时，浏览器通过 `POST /files/link/<workspace>/<name>`（令牌放在请求头中）打开文件，该请求返回一个仅针对该文件签名、有效期五分钟的链接。`GET /health` 报告房间数和通话数。

**请将它置于 TLS 之后。** 浏览器会拒绝 `https://` 页面发起的 `ws://` 连接，而通过明文 HTTP 传输的令牌或通话信令在网络上是可读的。任何能升级 WebSocket 的反向代理都可以（nginx、Caddy、Traefik）。Caddy 示例：

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

然后让应用指向 `wss://teams.example.com/sync` 和 `wss://teams.example.com/signal`。

**令牌是一个共享密钥**，保存在浏览器中：它可以把陌生人挡在小团队的服务器之外，但并不是按用户的身份验证。若需要真正的账户，请在前面放置你自己的身份验证（检查会话的代理，或将 `tokenOk` 替换为令牌校验），并保持相同的三个约定。

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
COPY server/index.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data
VOLUME /data
EXPOSE 8787
CMD ["node", "index.mjs"]
```

运行时请设置 `TEAMS_TOKEN` 和 `TEAMS_ORIGINS`，并置于你的 TLS 代理之后。

## 使用其他服务器

任何遵守核心 `docs/teams-area.md` 中约定的服务器都可以：同步使用标准的 `y-websocket` 服务器，信令使用任意 WebSocket 扇出，文件使用任意对象存储（给 `<teams-app>` 提供一个 `uploadFile` 函数，或在元素上设置 `uploadFile`，附件就会存放到那里）。
