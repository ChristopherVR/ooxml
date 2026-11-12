# OneDrive 链接与 Office 文档

如需分步设置（Microsoft Entra 注册、登录和完整示例），请从 [OneDrive 快速入门](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md)开始。

OpenTeams 可以携带指向 DOCX、XLSX、PPTX 或 VSDX 文件的 OneDrive 链接。集成的 Office 套件会通过对应的文档编辑器，以及共享的 `ooxml-core/embedding/onedrive` 连接打开该文件。OpenTeams 本身是一个工作空间，而不是第五个 Office 文件编辑器，也不是 Microsoft Teams。

同样的区别也适用于 `openteams-react-viewer`、`openteams-vue-viewer`、`openteams-angular-viewer`、`openteams-svelte-viewer`、`openteams-solid-viewer`、`openteams-vanilla-viewer` 以及 `ooxml-ui/teams` 元素：

| 共享项               | 编辑内容的保存位置                                                         |
| -------------------- | -------------------------------------------------------------------------- |
| OneDrive 链接        | 仅当经授权的宿主通过其 OneDrive 会话保存时，才保存到原始 OneDrive 项目中。 |
| 本地工作空间文件链接 | 保存在本地 Office 工作空间中；它不是跨设备的云端链接。                     |
| OpenTeams 服务器附件 | 保存在该服务器的存储中；上传的副本与 OneDrive 原件相互独立。               |

套件中已连接文件的“共享”流程会请求明确的查看或编辑角色以及链接范围。发布或复制已有项目的 URL 本身并不会授予他人访问权限。当租户拒绝受限链接时，切勿将匿名链接作为备选方案。`openteams-server` 不会获取 Microsoft 凭据，不会镜像云盘文件夹，也不会合并 Office 文件的更改。

有关身份验证、包引用和基于权限的共享，请使用 [Office OneDrive 集成指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams)。其[测试指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)区分了测试专用的传输方式与真实账户的验收。本项目未附带任何模拟界面或服务。

Yjs 聊天/在线状态以及文档协作房间与 Microsoft Office 的共同创作是相互独立的。把同一个 OneDrive 链接发送给两个人，并不会自动同步他们的编辑器。
