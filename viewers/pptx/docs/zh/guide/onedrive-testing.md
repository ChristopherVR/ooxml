---
title: 测试 OneDrive 集成
description: 借助测试专用的传输层和经授权的 Microsoft 账户验收，验证 Office 共享的 OneDrive 连接。
---

# 测试 OneDrive 集成

Word、Excel、PowerPoint 和 Visio 共用同一个 OneDrive 连接。其自动化传输层仅属于测试夹具。这里没有随产品发布的模拟界面、公开的模拟服务或 OneDrive 模拟演示，也无需配置或部署它们。

有关命令、各格式的检查项和真实账户验收矩阵，请参阅 [Office 测试指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)。[集成指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)则记录了包和框架的配置、身份验证、共享以及冲突恢复。

自动化回归测试无需 Microsoft 凭据即可验证 Graph 请求顺序和编辑器的保存连接。它们涵盖快照与文件标识、权限限制、修订条件、保留的本地编辑、独立的云端副本以及错误处理。但它们无法证明 Microsoft 登录、同意授权、条件访问、已部署源的 CORS 或原生 PowerPoint 的保真度。

Microsoft Graph Explorer 允许使用 Microsoft 现有的应用注册进行手动 API 检查。它不能验证应用自身的注册或登录流程。请使用一次性演示文稿并遵循测试指南；不要把 Graph Explorer 的客户端 ID 或借用的凭据当作应用的身份验证方案发布。

如需通过套件进行实际编辑，请在入口加载之前配置宿主的 Microsoft 令牌提供程序，然后使用 **Open from OneDrive**（从 OneDrive 打开）。真实的个人或工作账户以及 SharePoint 租户策略仍是必须完成的验收用例。任何无需注册的测试都不能作为实时 Microsoft 联合编辑或保证无冲突并发编辑的证据。
