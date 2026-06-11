# 打卡质检数据生成工具

纯前端打卡质检网页工具。Excel 文件只在浏览器本地处理，不会上传服务器。

[在线使用](https://checkin-quality-tool.vercel.app)

读取使用 SheetJS 0.20.3 的 dense 模式，并在 Web Worker 中运行；导出使用 fflate 流式生成标准 XLSX，避免大文件阻塞页面或一次性拼接超大 XML。

## 功能

- 上传课堂反馈名单和聊天质检明细
- 删除私聊、无邮箱、非员工发送及引用回复
- 按教师邮箱和学员关键词匹配打卡记录
- 英文匹配忽略大小写，异常单字姓名自动使用弱匹配
- 支持项目内置 `data/whitelist.csv` 白名单
- 导出打卡结果、匹配明细、清洗后聊天和处理说明
- 支持约 10 万行、50MB 以上的 Excel 文件

## 在线部署

仓库可直接导入 Vercel，无需构建命令、后端或数据库。

1. 在 Vercel 选择 `Add New Project`
2. 导入本 GitHub 仓库
3. Framework Preset 选择 `Other`
4. 保持 Build Command 为空
5. 点击 Deploy

`vercel.json` 已包含静态站点与安全响应头配置。

## 本地开发

Web Worker 需要通过 HTTP 访问。可在仓库目录运行：

```bash
python3 -m http.server 8765
```

然后访问：

```text
http://127.0.0.1:8765
```

## 输出工作表

- 打卡结果
- 匹配明细
- 清洗后聊天
- 处理说明
- 内置白名单

白名单优先按学员号关联。`处理方式`支持：

- `免检`：不再检查聊天；`发送情况`直接输出`已发送`，`匹配结论`保留`白名单免检`。
- `别名`：使用`匹配别名`的后两字补充匹配，聊天实际命中后才判已发送。

## 代码结构

- `worker.js`：文件读取、名单/聊天预处理和整体流程
- `modules/whitelist.js`：CSV白名单解析与学员关联
- `modules/matching.js`：教师邮箱、学员姓名和别名匹配

浏览器建议使用最新版 Chrome 或 Edge，并关闭不必要的标签页，以便为 50MB 以上 Excel 留出足够内存。

## 数据安全

- 文件由浏览器的 `File` API 读取
- 清洗、匹配和导出均在 Web Worker 中执行
- 项目没有文件上传接口、数据库或分析脚本
- 请勿将真实教师、学员或聊天数据提交到公开仓库

## 开源许可

[MIT](./LICENSE)
