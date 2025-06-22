class TableClipApp {
    constructor() {
        this.parser = new TableParser();
        this.exporter = null;
        this.currentFormat = 'html';
        this.currentTableData = null;
        
        this.initializeElements();
        this.bindEvents();
        this.checkClipboardPermission();
    }

    /**
     * 初始化DOM元素引用
     */
    initializeElements() {
        // 按钮元素
        this.pasteBtn = document.getElementById('pasteBtn');
        this.clearBtn = document.getElementById('clearBtn');
        this.copyOutputBtn = document.getElementById('copyOutput');
        this.downloadOutputBtn = document.getElementById('downloadOutput');
        
        // 格式按钮
        this.exportHtmlBtn = document.getElementById('exportHtml');
        this.exportMarkdownBtn = document.getElementById('exportMarkdown');
        this.exportTextileBtn = document.getElementById('exportTextile');
        this.exportRstBtn = document.getElementById('exportRst');
        
        // 选项复选框
        this.preserveMergedCb = document.getElementById('preserveMerged');
        this.includeStylesCb = document.getElementById('includeStyles');
        this.responsiveTableCb = document.getElementById('responsiveTable');
        
        // HTML样式选项
        this.htmlStylesPanel = document.getElementById('htmlStylesPanel');
        this.headerRowBoldCb = document.getElementById('headerRowBold');
        this.headerColBoldCb = document.getElementById('headerColBold');
        this.alternatingRowsCb = document.getElementById('alternatingRows');
        this.stylePresetSelect = document.getElementById('stylePreset');
        this.tableWidthSelect = document.getElementById('tableWidth');
        this.horizontalAlignSelect = document.getElementById('horizontalAlign');
        this.verticalAlignSelect = document.getElementById('verticalAlign');
        
        // 显示区域
        this.tablePreview = document.getElementById('tablePreview');
        this.outputCode = document.getElementById('outputCode');
        this.statusMessage = document.getElementById('statusMessage');
    }

    /**
     * 绑定事件监听器
     */
    bindEvents() {
        // 粘贴按钮
        this.pasteBtn.addEventListener('click', () => this.handlePaste());
        
        // 清空按钮
        this.clearBtn.addEventListener('click', () => this.handleClear());
        
        // 格式按钮
        this.exportHtmlBtn.addEventListener('click', () => this.handleFormatChange('html'));
        this.exportMarkdownBtn.addEventListener('click', () => this.handleFormatChange('markdown'));
        this.exportTextileBtn.addEventListener('click', () => this.handleFormatChange('textile'));
        this.exportRstBtn.addEventListener('click', () => this.handleFormatChange('rst'));
        
        // 选项变化
        this.preserveMergedCb.addEventListener('change', () => this.updateOutput());
        this.includeStylesCb.addEventListener('change', () => this.updateOutput());
        this.responsiveTableCb.addEventListener('change', () => this.updateOutput());
        
        // HTML样式选项变化
        this.headerRowBoldCb.addEventListener('change', () => this.updateOutput());
        this.headerColBoldCb.addEventListener('change', () => this.updateOutput());
        this.alternatingRowsCb.addEventListener('change', () => this.updateOutput());
        this.stylePresetSelect.addEventListener('change', () => this.handleStylePresetChange());
        this.tableWidthSelect.addEventListener('change', () => this.updateOutput());
        this.horizontalAlignSelect.addEventListener('change', () => this.updateOutput());
        this.verticalAlignSelect.addEventListener('change', () => this.updateOutput());
        
        // 输出操作
        this.copyOutputBtn.addEventListener('click', () => this.handleCopyOutput());
        this.downloadOutputBtn.addEventListener('click', () => this.handleDownloadOutput());
        
        // 键盘快捷键
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey || e.metaKey) {
                switch (e.key) {
                    case 'v':
                        if (e.target === document.body) {
                            e.preventDefault();
                            this.handlePaste();
                        }
                        break;
                    case 'c':
                        if (e.shiftKey && this.outputCode.textContent.trim()) {
                            e.preventDefault();
                            this.handleCopyOutput();
                        }
                        break;
                }
            }
        });
    }

    /**
     * 检查剪贴板权限
     */
    async checkClipboardPermission() {
        try {
            const permission = await navigator.permissions.query({name: 'clipboard-read'});
            if (permission.state === 'denied') {
                this.showStatus('剪贴板访问被拒绝，请在浏览器设置中允许访问', 'warning');
            }
        } catch (error) {
            console.log('无法检查剪贴板权限:', error);
        }
    }

    /**
     * 处理粘贴操作
     */
    async handlePaste() {
        try {
            this.pasteBtn.disabled = true;
            this.pasteBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 解析中...';
            
            const tableData = await this.parser.parseFromClipboard();
            this.currentTableData = tableData;
            this.exporter = new FormatExporters(tableData);
            
            this.updatePreview();
            this.updateOutput();
            
            this.showStatus('表格解析成功！', 'success');
            
        } catch (error) {
            console.error('粘贴失败:', error);
            this.showStatus(`粘贴失败: ${error.message}`, 'error');
        } finally {
            this.pasteBtn.disabled = false;
            this.pasteBtn.innerHTML = '<i class="fas fa-paste"></i> 粘贴表格数据';
        }
    }

    /**
     * 处理清空操作
     */
    handleClear() {
        this.currentTableData = null;
        this.exporter = null;
        this.parser.tableData = null;
        
        this.tablePreview.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-table"></i>
                <p>暂无表格数据</p>
                <p class="hint">请从 Excel 复制表格并粘贴</p>
            </div>
        `;
        
        this.outputCode.textContent = '<!-- 转换后的代码将显示在这里 -->';
        
        this.showStatus('已清空所有数据', 'success');
    }

    /**
     * 处理格式切换
     */
    handleFormatChange(format) {
        // 更新按钮状态
        document.querySelectorAll('.btn-format').forEach(btn => {
            btn.classList.remove('active');
        });
        
        const formatMap = {
            'html': this.exportHtmlBtn,
            'markdown': this.exportMarkdownBtn,
            'textile': this.exportTextileBtn,
            'rst': this.exportRstBtn
        };
        
        if (formatMap[format]) {
            formatMap[format].classList.add('active');
        }
        
        this.currentFormat = format;
        
        // 显示或隐藏HTML样式面板
        if (format === 'html') {
            this.htmlStylesPanel.style.display = 'block';
        } else {
            this.htmlStylesPanel.style.display = 'none';
        }
        
        this.updateOutput();
    }

    /**
     * 更新预览区域
     */
    updatePreview() {
        if (!this.currentTableData) {
            return;
        }

        const options = this.getExportOptions();
        const previewHtml = this.parser.generatePreviewHtml(options);
        this.tablePreview.innerHTML = previewHtml;
    }

    /**
     * 更新输出区域
     */
    updateOutput() {
        if (!this.exporter || !this.currentTableData) {
            return;
        }

        const options = this.getExportOptions();
        let output = '';

        try {
            switch (this.currentFormat) {
                case 'html':
                    output = this.exporter.exportToHtml(options);
                    break;
                case 'markdown':
                    output = this.exporter.exportToMarkdown(options);
                    break;
                case 'textile':
                    output = this.exporter.exportToTextile(options);
                    break;
                case 'rst':
                    output = this.exporter.exportToRst(options);
                    break;
                default:
                    output = '不支持的格式';
            }

            this.outputCode.textContent = output;
        } catch (error) {
            console.error('输出生成失败:', error);
            this.outputCode.textContent = `错误: ${error.message}`;
        }
    }

    /**
     * 获取导出选项
     */
    getExportOptions() {
        const options = {
            preserveMerged: this.preserveMergedCb.checked,
            includeStyles: this.includeStylesCb.checked,
            responsiveTable: this.responsiveTableCb.checked,
            standalone: true
        };
        
        // 添加HTML特定选项
        if (this.currentFormat === 'html') {
            options.htmlStyles = {
                headerRowBold: this.headerRowBoldCb.checked,
                headerColBold: this.headerColBoldCb.checked,
                alternatingRows: this.alternatingRowsCb.checked,
                preset: this.stylePresetSelect.value
            };
            options.tableWidth = this.tableWidthSelect.value;
            options.horizontalAlign = this.horizontalAlignSelect.value;
            options.verticalAlign = this.verticalAlignSelect.value;
        }
        
        return options;
    }

    /**
     * 处理复制输出
     */
    async handleCopyOutput() {
        const content = this.outputCode.textContent;
        
        if (!content || content.trim() === '<!-- 转换后的代码将显示在这里 -->') {
            this.showStatus('没有可复制的内容', 'warning');
            return;
        }

        try {
            await navigator.clipboard.writeText(content);
            this.showStatus('代码已复制到剪贴板', 'success');
        } catch (error) {
            // 回退方案
            try {
                this.fallbackCopyToClipboard(content);
                this.showStatus('代码已复制到剪贴板', 'success');
            } catch (fallbackError) {
                console.error('复制失败:', fallbackError);
                this.showStatus('复制失败，请手动选择复制', 'error');
            }
        }
    }

    /**
     * 回退复制方法
     */
    fallbackCopyToClipboard(text) {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        
        try {
            const result = document.execCommand('copy');
            if (!result) {
                throw new Error('execCommand返回false');
            }
        } finally {
            document.body.removeChild(textArea);
        }
    }

    /**
     * 处理下载输出
     */
    handleDownloadOutput() {
        const content = this.outputCode.textContent;
        
        if (!content || content.trim() === '<!-- 转换后的代码将显示在这里 -->') {
            this.showStatus('没有可下载的内容', 'warning');
            return;
        }

        try {
            const extension = this.exporter.getFileExtension(this.currentFormat);
            const mimeType = this.exporter.getMimeType(this.currentFormat);
            const filename = `table_export_${Date.now()}${extension}`;

            const blob = new Blob([content], { type: mimeType });
            const url = URL.createObjectURL(blob);

            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            link.style.display = 'none';

            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            // 清理URL对象
            setTimeout(() => URL.revokeObjectURL(url), 100);

            this.showStatus(`文件已下载: ${filename}`, 'success');
        } catch (error) {
            console.error('下载失败:', error);
            this.showStatus('下载失败，请重试', 'error');
        }
    }

    /**
     * 显示状态消息
     */
    showStatus(message, type = 'success') {
        this.statusMessage.textContent = message;
        this.statusMessage.className = `status-message ${type}`;
        this.statusMessage.classList.add('show');

        // 自动隐藏
        setTimeout(() => {
            this.statusMessage.classList.remove('show');
        }, 3000);
    }

    /**
     * 处理样式预设变化
     */
    handleStylePresetChange() {
        const preset = this.stylePresetSelect.value;
        
        switch (preset) {
            case 'default':
                this.headerRowBoldCb.checked = true;
                this.headerColBoldCb.checked = false;
                this.alternatingRowsCb.checked = true;
                break;
            case 'minimal':
                this.headerRowBoldCb.checked = false;
                this.headerColBoldCb.checked = false;
                this.alternatingRowsCb.checked = false;
                break;
            case 'professional':
                this.headerRowBoldCb.checked = true;
                this.headerColBoldCb.checked = true;
                this.alternatingRowsCb.checked = true;
                break;
            case 'modern':
                this.headerRowBoldCb.checked = true;
                this.headerColBoldCb.checked = false;
                this.alternatingRowsCb.checked = true;
                break;
            case 'custom':
                // 保持当前设置不变
                break;
        }
        
        this.updateOutput();
    }

    /**
     * 获取表格统计信息
     */
    getTableStats() {
        if (!this.currentTableData) {
            return null;
        }

        const stats = {
            rows: this.currentTableData.rows.length,
            columns: this.currentTableData.maxColumns,
            hasHeader: this.currentTableData.hasHeader,
            mergedCells: 0
        };

        // 计算合并单元格数量
        this.currentTableData.rows.forEach(row => {
            row.cells.forEach(cell => {
                if (cell && (cell.colspan > 1 || cell.rowspan > 1)) {
                    stats.mergedCells++;
                }
            });
        });

        return stats;
    }
}

// 为按钮添加激活状态样式
const style = document.createElement('style');
style.textContent = `
.btn-format.active {
    background: #667eea !important;
    color: white !important;
    border-color: #667eea !important;
}
`;
document.head.appendChild(style);

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    const app = new TableClipApp();
    
    // 设置默认格式
    app.handleFormatChange('html');
    
    // 全局错误处理
    window.addEventListener('error', (event) => {
        console.error('全局错误:', event.error);
    });
    
    window.addEventListener('unhandledrejection', (event) => {
        console.error('未处理的Promise拒绝:', event.reason);
    });
}); 