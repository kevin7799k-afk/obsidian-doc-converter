# 📄 Obsidian Doc Converter

一个 Obsidian 插件，自动检测并转换 Office 文档为 Markdown，支持侧边栏预览。

## ✨ 功能

- 🔄 **自动转换**：拖入 Office 文件到 vault 时自动触发转换
- 👁️ **侧边栏预览**：转换后在右侧面板实时预览内容
- 📝 **保存 Markdown**：转换结果自动保存为 .md 文件
- 🖱️ **右键菜单**：右键任意 Office 文件即可转换或预览
- 📦 **批量转换**：一键转换 vault 中所有 Office 文档

## 📁 支持格式

| 格式 | 说明 | 支持度 |
|------|------|--------|
| .docx | Word 文档（新版）| ✅ 完整支持 |
| .xlsx | Excel 表格 | ✅ 完整支持 |
| .pptx | PowerPoint 演示文稿 | ✅ 文字提取 |
| .doc  | Word 文档（旧版）| ⚠️ 基本支持 |

## 🚀 安装方法

### 方法一：使用 BRAT（推荐，可自动更新）

1. 在 Obsidian 社区插件中安装并启用 **BRAT**
2. 打开 BRAT 设置，点击 **Add Beta plugin**
3. 输入本仓库地址：`https://github.com/kevin7799k-afk/obsidian-doc-converter`
4. 点击 **Add Plugin**，BRAT 自动安装
5. 前往 **设置 → 第三方插件** → 启用 **Doc Converter**

### 方法二：手动安装

1. 从 [Releases](https://github.com/kevin7799k-afk/obsidian-doc-converter/releases) 下载 `main.js`、`manifest.json`、`styles.css`
2. 在 vault 中创建文件夹：`.obsidian/plugins/doc-converter/`
3. 将三个文件放入该文件夹
4. 重启 Obsidian，在第三方插件中启用

## 📖 使用方法

| 操作 | 方式 |
|------|------|
| 自动转换 | 将 Office 文件拖入 vault 自动触发 |
| 手动转换 | 右键文件 → **📄 转换为 Markdown** |
| 仅预览 | 右键文件 → **👁️ 预览内容** |
| 批量转换 | `Ctrl+P` → 搜索「批量转换」|
| 打开预览面板 | 点击左侧 ribbon 文件图标 |

## ⚙️ 设置选项

- **自动转换**：新文件加入 vault 时自动转换（默认开启）
- **保存 Markdown**：转换结果保存为 .md 文件（默认开启）
- **自动打开预览**：转换后自动显示侧边栏预览（默认开启）
- **输出文件夹**：指定 Markdown 保存位置（默认与原文件同目录）

## ⚠️ 注意事项

- 仅支持桌面版 Obsidian（Windows / macOS / Linux）
- `.doc` 旧版格式解析成功率较低，建议先用 Word 另存为 `.docx`
- 首次使用时会从 CDN 加载解析库，需要网络连接

## 📜 License

MIT
