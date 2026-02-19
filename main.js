/*
 * Obsidian Doc Converter Plugin
 * 支持格式：.docx, .xlsx, .doc, .pptx
 * 功能：自动检测并转换为 Markdown，同时支持侧边栏预览
 */

'use strict';

const { Plugin, PluginSettingTab, Setting, Notice, ItemView, WorkspaceLeaf, TFile, Modal } = require('obsidian');

async function loadMammoth() {
    if (window._mammoth) return window._mammoth;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js';
        script.onload = () => { window._mammoth = window.mammoth; resolve(window.mammoth); };
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

async function loadXLSX() {
    if (window._XLSX) return window._XLSX;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
        script.onload = () => { window._XLSX = window.XLSX; resolve(window.XLSX); };
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

async function convertDocx(arrayBuffer) {
    const mammoth = await loadMammoth();
    const result = await mammoth.convertToMarkdown({ arrayBuffer });
    return result.value;
}

async function convertXlsx(arrayBuffer) {
    const XLSX = await loadXLSX();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    let markdown = '';
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!data || data.length === 0) continue;
        markdown += `\n## 📊 ${sheetName}\n\n`;
        const maxCols = Math.max(...data.map(r => r.length));
        if (maxCols === 0) continue;
        const header = data[0];
        markdown += '| ' + Array.from({ length: maxCols }, (_, i) => String(header[i] ?? '').replace(/\|/g, '\\|') || `列${i + 1}`).join(' | ') + ' |\n';
        markdown += '| ' + Array(maxCols).fill('---').join(' | ') + ' |\n';
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            markdown += '| ' + Array.from({ length: maxCols }, (_, j) => String(row[j] ?? '').replace(/\|/g, '\\|').replace(/\n/g, ' ')).join(' | ') + ' |\n';
        }
        markdown += '\n';
    }
    return markdown;
}

async function convertPptx(arrayBuffer) {
    try {
        const uint8 = new Uint8Array(arrayBuffer);
        const text = new TextDecoder('utf-8', { fatal: false }).decode(uint8);
        const matches = [...text.matchAll(/<a:t[^>]*>([^<]+)<\/a:t>/g)];
        if (matches.length === 0) return '> ⚠️ 无法提取 PPTX 文字内容，文件可能加密或格式不支持。';
        let markdown = '# 📑 演示文稿内容\n\n';
        let chunk = [], page = 1;
        for (const m of matches) {
            const t = m[1].trim();
            if (!t) continue;
            chunk.push(t);
            if (chunk.length >= 30) {
                markdown += `## 幻灯片 ${page}\n\n` + chunk.join('\n\n') + '\n\n';
                chunk = []; page++;
            }
        }
        if (chunk.length > 0) markdown += `## 幻灯片 ${page}\n\n` + chunk.join('\n\n') + '\n\n';
        return markdown;
    } catch (e) {
        return `> ⚠️ PPTX 解析失败：${e.message}`;
    }
}

async function convertDoc(arrayBuffer) {
    try {
        const mammoth = await loadMammoth();
        const result = await mammoth.convertToMarkdown({ arrayBuffer });
        if (result.value && result.value.trim().length > 0) return result.value;
        return '> ⚠️ .doc 旧版格式解析内容为空，建议先用 Word 另存为 .docx 再转换。';
    } catch (e) {
        return `> ⚠️ .doc 旧版格式解析失败：${e.message}\n> 建议先用 Word 另存为 .docx 再转换。`;
    }
}

async function convertFile(file, app) {
    const ext = file.extension.toLowerCase();
    const arrayBuffer = await app.vault.readBinary(file);
    switch (ext) {
        case 'docx': return await convertDocx(arrayBuffer);
        case 'xlsx': return await convertXlsx(arrayBuffer);
        case 'pptx': return await convertPptx(arrayBuffer);
        case 'doc':  return await convertDoc(arrayBuffer);
        default: throw new Error(`不支持的格式：.${ext}`);
    }
}

const VIEW_TYPE = 'doc-converter-preview';

class DocPreviewView extends ItemView {
    constructor(leaf) { super(leaf); this.content = ''; this.title = '文档预览'; }
    getViewType() { return VIEW_TYPE; }
    getDisplayText() { return this.title; }
    getIcon() { return 'file-text'; }
    async onOpen() { this.render(); }
    setContent(title, markdown) { this.title = title; this.content = markdown; this.render(); }
    render() {
        const container = this.containerEl.children[1];
        container.empty();
        container.style.cssText = 'padding:16px;overflow:auto;height:100%;';
        if (!this.content) {
            container.createEl('div', { text: '📂 选择文件后，预览将显示在这里', attr: { style: 'color:var(--text-muted);text-align:center;margin-top:40px;' } });
            return;
        }
        const el = container.createEl('div', { cls: 'markdown-rendered doc-converter-preview' });
        try {
            const { MarkdownRenderer } = require('obsidian');
            MarkdownRenderer.renderMarkdown(this.content, el, '', this);
        } catch (e) {
            el.innerHTML = `<pre style="white-space:pre-wrap">${this.content}</pre>`;
        }
    }
}

const DEFAULT_SETTINGS = { autoConvert: true, outputFolder: '', openPreview: true, saveMarkdown: true };

class DocConverterSettingTab extends PluginSettingTab {
    constructor(app, plugin) { super(app, plugin); this.plugin = plugin; }
    display() {
        const { containerEl } = this;
        containerEl.empty();
        containerEl.createEl('h2', { text: '📄 文档转换器设置' });
        new Setting(containerEl).setName('自动转换').setDesc('将新添加的 Office 文件自动转换为 Markdown').addToggle(t => t.setValue(this.plugin.settings.autoConvert).onChange(async v => { this.plugin.settings.autoConvert = v; await this.plugin.saveSettings(); }));
        new Setting(containerEl).setName('保存 Markdown 文件').setDesc('转换后将 Markdown 保存到 vault 中').addToggle(t => t.setValue(this.plugin.settings.saveMarkdown).onChange(async v => { this.plugin.settings.saveMarkdown = v; await this.plugin.saveSettings(); }));
        new Setting(containerEl).setName('自动打开预览').setDesc('转换后自动在侧边栏显示预览').addToggle(t => t.setValue(this.plugin.settings.openPreview).onChange(async v => { this.plugin.settings.openPreview = v; await this.plugin.saveSettings(); }));
        new Setting(containerEl).setName('输出文件夹').setDesc('Markdown 文件保存位置（留空表示与原文件同目录）').addText(t => t.setPlaceholder('例如：converted/').setValue(this.plugin.settings.outputFolder).onChange(async v => { this.plugin.settings.outputFolder = v; await this.plugin.saveSettings(); }));
    }
}

class DocConverterPlugin extends Plugin {
    async onload() {
        await this.loadSettings();
        this.registerView(VIEW_TYPE, leaf => new DocPreviewView(leaf));
        for (const ext of ['docx', 'xlsx', 'doc', 'pptx']) {
            try { this.registerExtensions([ext], 'doc-converter'); } catch (e) {}
        }
        this.registerEvent(this.app.workspace.on('file-menu', (menu, file) => {
            if (!(file instanceof TFile)) return;
            const ext = file.extension.toLowerCase();
            if (!['docx', 'xlsx', 'doc', 'pptx'].includes(ext)) return;
            menu.addItem(item => item.setTitle('📄 转换为 Markdown').setIcon('file-text').onClick(() => this.convertAndHandle(file)));
            menu.addItem(item => item.setTitle('👁️ 预览内容').setIcon('eye').onClick(() => this.convertAndPreview(file)));
        }));
        this.addCommand({ id: 'convert-active-doc', name: '转换当前文档', callback: () => { const file = this.app.workspace.getActiveFile(); if (!file) return new Notice('❌ 没有选中的文件'); this.convertAndHandle(file); } });
        this.addCommand({ id: 'convert-all-docs', name: '批量转换 vault 中所有 Office 文档', callback: () => this.batchConvert() });
        this.addCommand({ id: 'open-preview', name: '打开文档预览面板', callback: () => this.activatePreviewView() });
        if (this.settings.autoConvert) {
            this.registerEvent(this.app.vault.on('create', file => {
                if (!(file instanceof TFile)) return;
                if (!['docx', 'xlsx', 'doc', 'pptx'].includes(file.extension.toLowerCase())) return;
                setTimeout(() => this.convertAndHandle(file), 500);
            }));
        }
        this.addSettingTab(new DocConverterSettingTab(this.app, this));
        this.addRibbonIcon('file-text', '文档转换器', () => this.activatePreviewView());
        console.log('✅ Doc Converter Plugin loaded');
    }

    async convertAndHandle(file) {
        const notice = new Notice(`⏳ 正在转换 ${file.name}...`, 0);
        try {
            const markdown = await convertFile(file, this.app);
            notice.hide();
            if (this.settings.saveMarkdown) await this.saveMarkdown(file, markdown);
            if (this.settings.openPreview) await this.showPreview(file.basename, markdown);
            new Notice(`✅ ${file.name} 转换完成！`, 4000);
        } catch (e) { notice.hide(); new Notice(`❌ 转换失败：${e.message}`, 6000); console.error(e); }
    }

    async convertAndPreview(file) {
        const notice = new Notice(`⏳ 正在解析 ${file.name}...`, 0);
        try { const markdown = await convertFile(file, this.app); notice.hide(); await this.showPreview(file.basename, markdown); }
        catch (e) { notice.hide(); new Notice(`❌ 解析失败：${e.message}`, 6000); }
    }

    async saveMarkdown(file, markdown) {
        let folder = this.settings.outputFolder.trim();
        if (folder && !folder.endsWith('/')) folder += '/';
        const basePath = folder ? folder + file.basename + '.md' : file.parent.path + (file.parent.path ? '/' : '') + file.basename + '.md';
        const header = `---\ntitle: ${file.basename}\nsource: ${file.path}\nconverted: ${new Date().toISOString()}\n---\n\n`;
        const fullContent = header + markdown;
        const existing = this.app.vault.getAbstractFileByPath(basePath);
        if (existing instanceof TFile) { await this.app.vault.modify(existing, fullContent); }
        else {
            if (folder) { try { await this.app.vault.createFolder(folder.slice(0, -1)); } catch (e) {} }
            await this.app.vault.create(basePath, fullContent);
        }
        const mdFile = this.app.vault.getAbstractFileByPath(basePath);
        if (mdFile instanceof TFile) await this.app.workspace.getLeaf(false).openFile(mdFile);
    }

    async showPreview(title, markdown) { await this.activatePreviewView(); const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE); if (leaves.length > 0) leaves[0].view.setContent(title, markdown); }

    async activatePreviewView() {
        const existing = this.app.workspace.getLeavesOfType(VIEW_TYPE);
        if (existing.length > 0) { this.app.workspace.revealLeaf(existing[0]); return; }
        const leaf = this.app.workspace.getRightLeaf(false);
        await leaf.setViewState({ type: VIEW_TYPE, active: true });
        this.app.workspace.revealLeaf(leaf);
    }

    async batchConvert() {
        const files = this.app.vault.getFiles().filter(f => ['docx', 'xlsx', 'doc', 'pptx'].includes(f.extension.toLowerCase()));
        if (files.length === 0) { new Notice('📂 vault 中没有找到 Office 文件'); return; }
        new Notice(`🔄 开始批量转换 ${files.length} 个文件...`);
        let success = 0, fail = 0;
        for (const file of files) {
            try { const markdown = await convertFile(file, this.app); if (this.settings.saveMarkdown) await this.saveMarkdown(file, markdown); success++; }
            catch (e) { fail++; console.error(`转换 ${file.name} 失败:`, e); }
        }
        new Notice(`✅ 批量转换完成：成功 ${success} 个，失败 ${fail} 个`, 6000);
    }

    async loadSettings() { this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData()); }
    async saveSettings() { await this.saveData(this.settings); }
    onunload() { this.app.workspace.detachLeavesOfType(VIEW_TYPE); }
}

module.exports = DocConverterPlugin;