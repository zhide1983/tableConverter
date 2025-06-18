class FormatExporters {
    constructor(tableData) {
        this.tableData = tableData;
    }

    /**
     * 导出为HTML格式
     */
    exportToHtml(options = {}) {
        const {
            preserveMerged = true,
            includeStyles = true,
            responsiveTable = false,
            standalone = true,
            htmlStyles = {}
        } = options;

        if (!this.tableData || !this.tableData.rows) {
            return '';
        }

        let html = '';
        
        if (standalone) {
            html += '<!DOCTYPE html>\n<html>\n<head>\n';
            html += '<meta charset="UTF-8">\n';
            html += '<title>导出的表格</title>\n';
            html += '<style>\n';
            html += this.getEnhancedTableCSS(responsiveTable, htmlStyles);
            html += '</style>\n';
            html += '</head>\n<body>\n';
        }

        const tableClass = responsiveTable ? ' class="responsive-table"' : '';
        html += `<table${tableClass}>\n`;

        // 生成表头
        if (this.tableData.hasHeader) {
            html += '  <thead>\n';
            html += this.generateHtmlRow(this.tableData.rows[0], preserveMerged, includeStyles, 2, htmlStyles, 'header');
            html += '  </thead>\n';
        }

        // 生成表体
        const startRow = this.tableData.hasHeader ? 1 : 0;
        if (startRow < this.tableData.rows.length) {
            html += '  <tbody>\n';
            for (let i = startRow; i < this.tableData.rows.length; i++) {
                html += this.generateHtmlRow(this.tableData.rows[i], preserveMerged, includeStyles, 2, htmlStyles, 'body', i - startRow);
            }
            html += '  </tbody>\n';
        }

        html += '</table>';

        if (standalone) {
            html += '\n</body>\n</html>';
        }

        return html;
    }

    /**
     * 生成HTML行
     */
    generateHtmlRow(row, preserveMerged, includeStyles, indent = 0, htmlStyles = {}, rowType = 'body', rowIndex = 0) {
        const spaces = ' '.repeat(indent);
        let rowClass = '';
        
        // 添加阴阳行样式
        if (htmlStyles.alternatingRows && rowType === 'body') {
            rowClass = rowIndex % 2 === 0 ? ' class="even-row"' : ' class="odd-row"';
        }
        
        let html = `${spaces}<tr${rowClass}>\n`;

        for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
            const cell = row.cells[colIndex];
            if (!cell || cell.isMerged) continue;

            const tag = cell.type === 'th' ? 'th' : 'td';
            let cellHtml = `${spaces}  <${tag}`;
            let cellClass = '';

            // 添加列样式类
            if (htmlStyles.headerColBold && colIndex === 0) {
                cellClass = ' class="header-col"';
            }

            if (cellClass) {
                cellHtml += cellClass;
            }

            // 添加合并属性
            if (preserveMerged) {
                if (cell.colspan > 1) cellHtml += ` colspan="${cell.colspan}"`;
                if (cell.rowspan > 1) cellHtml += ` rowspan="${cell.rowspan}"`;
            }

            // 添加原始样式
            if (includeStyles && cell.styles && Object.keys(cell.styles).length > 0) {
                const styleStr = Object.entries(cell.styles)
                    .map(([key, value]) => `${this.camelToKebab(key)}: ${value}`)
                    .join('; ');
                cellHtml += ` style="${styleStr}"`;
            }

            cellHtml += `>${this.escapeHtml(cell.content)}</${tag}>\n`;
            html += cellHtml;
        }

        html += `${spaces}</tr>\n`;
        return html;
    }

    /**
     * 导出为Markdown格式
     */
    exportToMarkdown(options = {}) {
        const { preserveMerged = false } = options;

        if (!this.tableData || !this.tableData.rows.length) {
            return '';
        }

        // 计算每列的最大宽度
        const columnWidths = this.calculateMarkdownColumnWidths(preserveMerged);
        
        let markdown = '';
        let startRow = 0;

        // 处理表头
        if (this.tableData.hasHeader) {
            const headerRow = this.tableData.rows[0];
            markdown += '|';
            
            for (let colIndex = 0; colIndex < headerRow.cells.length; colIndex++) {
                const cell = headerRow.cells[colIndex];
                
                if (!cell) {
                    const width = columnWidths[colIndex] || 3;
                    markdown += ` ${' '.repeat(Math.max(0, width - 2))} |`;
                    continue;
                }
                
                if (cell.isMerged) {
                    const width = columnWidths[colIndex] || 3;
                    markdown += ` ${' '.repeat(Math.max(0, width - 2))} |`;
                    continue;
                }
                
                let content = cell.content || '';
                if (preserveMerged && (cell.colspan > 1 || cell.rowspan > 1)) {
                    content += ` (${cell.colspan}×${cell.rowspan})`;
                }
                
                const width = columnWidths[colIndex] || 3;
                const padding = Math.max(0, width - this.getDisplayWidth(content) - 2);
                const leftPad = Math.floor(padding / 2);
                const rightPad = padding - leftPad;
                
                markdown += ` ${' '.repeat(leftPad)}${this.escapeMarkdown(content)}${' '.repeat(rightPad)} |`;
            }
            
            markdown += '\n|';
            for (let colIndex = 0; colIndex < headerRow.cells.length; colIndex++) {
                const width = columnWidths[colIndex] || 3;
                const dashes = '-'.repeat(width);
                markdown += ` ${dashes} |`;
            }
            markdown += '\n';
            startRow = 1;
        }

        // 处理数据行
        for (let i = startRow; i < this.tableData.rows.length; i++) {
            const row = this.tableData.rows[i];
            markdown += '|';
            
            for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
                const cell = row.cells[colIndex];
                
                if (!cell) {
                    const width = columnWidths[colIndex] || 3;
                    markdown += ` ${' '.repeat(Math.max(0, width - 2))} |`;
                    continue;
                }
                
                if (cell.isMerged) {
                    const width = columnWidths[colIndex] || 3;
                    markdown += ` ${' '.repeat(Math.max(0, width - 2))} |`;
                    continue;
                }
                
                let content = cell.content || '';
                if (preserveMerged && (cell.colspan > 1 || cell.rowspan > 1)) {
                    content += ` (${cell.colspan}×${cell.rowspan})`;
                }
                
                const width = columnWidths[colIndex] || 3;
                const padding = Math.max(0, width - this.getDisplayWidth(content) - 2);
                const leftPad = Math.floor(padding / 2);
                const rightPad = padding - leftPad;
                
                markdown += ` ${' '.repeat(leftPad)}${this.escapeMarkdown(content)}${' '.repeat(rightPad)} |`;
            }
            
            markdown += '\n';
        }

        return markdown;
    }

    /**
     * 计算Markdown表格每列的显示宽度
     */
    calculateMarkdownColumnWidths(preserveMerged = false) {
        const widths = [];

        for (let colIndex = 0; colIndex < this.tableData.maxColumns; colIndex++) {
            let maxWidth = 3; // 最小宽度

            for (const row of this.tableData.rows) {
                if (colIndex < row.cells.length && row.cells[colIndex] && !row.cells[colIndex].isMerged) {
                    const cell = row.cells[colIndex];
                    let content = cell.content || '';
                    
                    // 如果有合并单元格，添加标记的长度
                    if (preserveMerged && (cell.colspan > 1 || cell.rowspan > 1)) {
                        content += ` (${cell.colspan}×${cell.rowspan})`;
                    }
                    
                    const displayWidth = this.getDisplayWidth(content);
                    maxWidth = Math.max(maxWidth, displayWidth + 2); // 加上左右空格
                }
            }

            widths.push(maxWidth);
        }

        return widths;
    }

    /**
     * 计算字符串的显示宽度（考虑中文字符）
     */
    getDisplayWidth(text) {
        if (!text) return 0;
        
        let width = 0;
        for (let i = 0; i < text.length; i++) {
            const char = text.charCodeAt(i);
            // 中文字符、日文字符、韩文字符等通常占2个显示宽度
            if ((char >= 0x4e00 && char <= 0x9fff) ||  // 中文
                (char >= 0x3400 && char <= 0x4dbf) ||  // 中文扩展A
                (char >= 0x20000 && char <= 0x2a6df) || // 中文扩展B
                (char >= 0x3040 && char <= 0x309f) ||  // 日文平假名
                (char >= 0x30a0 && char <= 0x30ff) ||  // 日文片假名
                (char >= 0xac00 && char <= 0xd7af)) {  // 韩文
                width += 2;
            } else {
                width += 1;
            }
        }
        return width;
    }

    /**
     * 导出为Textile格式
     */
    exportToTextile(options = {}) {
        const { preserveMerged = true } = options;

        if (!this.tableData || !this.tableData.rows.length) {
            return '';
        }

        let textile = '';

        for (let rowIndex = 0; rowIndex < this.tableData.rows.length; rowIndex++) {
            const row = this.tableData.rows[rowIndex];
            const isHeader = this.tableData.hasHeader && rowIndex === 0;

            for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
                const cell = row.cells[colIndex];
                if (!cell || cell.isMerged) continue;

                let cellMarkup = '';

                // 表头标记
                if (isHeader || cell.type === 'th') {
                    cellMarkup += '_. ';
                }

                // 合并单元格标记
                if (preserveMerged) {
                    if (cell.colspan > 1) {
                        cellMarkup += `\\${cell.colspan}. `;
                    }
                    if (cell.rowspan > 1) {
                        cellMarkup += `/${cell.rowspan}. `;
                    }
                }

                // 样式标记
                if (cell.styles) {
                    const styles = [];
                    if (cell.styles.textAlign === 'center') styles.push('=');
                    else if (cell.styles.textAlign === 'right') styles.push('>');
                    else if (cell.styles.textAlign === 'left') styles.push('<');
                    
                    if (styles.length > 0) {
                        cellMarkup += `${styles.join('')}. `;
                    }
                }

                cellMarkup += this.escapeTextile(cell.content);
                textile += `|${cellMarkup}`;
            }

            textile += '|\n';
        }

        return textile;
    }

    /**
     * 导出为reStructuredText格式
     */
    exportToRst(options = {}) {
        const { preserveMerged = true } = options;

        if (!this.tableData || !this.tableData.rows.length) {
            return '';
        }

        // 计算每列的最大宽度
        const columnWidths = this.calculateColumnWidths();
        
        let rst = '';

        // 生成顶部边框
        rst += this.generateRstBorder(columnWidths, '+', '=') + '\n';

        // 生成表头
        if (this.tableData.hasHeader) {
            rst += this.generateRstRow(this.tableData.rows[0], columnWidths, preserveMerged);
            rst += this.generateRstBorder(columnWidths, '+', '=') + '\n';
        }

        // 生成数据行
        const startRow = this.tableData.hasHeader ? 1 : 0;
        for (let i = startRow; i < this.tableData.rows.length; i++) {
            rst += this.generateRstRow(this.tableData.rows[i], columnWidths, preserveMerged);
            
            // 行间分隔符
            if (i < this.tableData.rows.length - 1) {
                rst += this.generateRstBorder(columnWidths, '+', '-') + '\n';
            }
        }

        // 生成底部边框
        rst += this.generateRstBorder(columnWidths, '+', '=') + '\n';

        return rst;
    }

    /**
     * 计算RST表格列宽
     */
    calculateColumnWidths() {
        const widths = [];

        for (let colIndex = 0; colIndex < this.tableData.maxColumns; colIndex++) {
            let maxWidth = 5; // 最小宽度，为合并单元格标记留空间

            for (const row of this.tableData.rows) {
                if (colIndex < row.cells.length && row.cells[colIndex] && !row.cells[colIndex].isMerged) {
                    const cell = row.cells[colIndex];
                    let content = cell.content || '';
                    
                    // 如果有合并单元格，添加标记的长度
                    if (cell.colspan > 1 || cell.rowspan > 1) {
                        content += ` [${cell.colspan}×${cell.rowspan}]`;
                    }
                    
                    maxWidth = Math.max(maxWidth, content.length + 2); // 加上左右空格
                }
            }

            widths.push(maxWidth);
        }

        return widths;
    }

    /**
     * 生成RST边框
     */
    generateRstBorder(columnWidths, corner, line) {
        let border = corner;
        for (const width of columnWidths) {
            border += line.repeat(width) + corner;
        }
        return border;
    }

    /**
     * 生成RST行
     */
    generateRstRow(row, columnWidths, preserveMerged) {
        let rst = '|';

        for (let colIndex = 0; colIndex < columnWidths.length; colIndex++) {
            const cell = (colIndex < row.cells.length) ? row.cells[colIndex] : null;
            
            if (!cell) {
                rst += ' '.repeat(columnWidths[colIndex]) + '|';
                continue;
            }

            if (cell.isMerged) {
                rst += ' '.repeat(columnWidths[colIndex]) + '|'; // 合并单元格的其他位置留空
                continue;
            }

            let content = cell.content || '';
            
            // 添加合并信息（作为注释）
            if (preserveMerged && (cell.colspan > 1 || cell.rowspan > 1)) {
                content += ` [${cell.colspan}×${cell.rowspan}]`;
            }

            // 确保内容长度不超过列宽
            if (content.length > columnWidths[colIndex] - 2) {
                content = content.substring(0, columnWidths[colIndex] - 5) + '...';
            }

            const padding = columnWidths[colIndex] - content.length;
            const leftPad = Math.floor(padding / 2);
            const rightPad = padding - leftPad;

            rst += ' '.repeat(leftPad) + content + ' '.repeat(rightPad) + '|';
        }

        return rst + '\n';
    }

    /**
     * 获取增强的表格CSS样式
     */
    getEnhancedTableCSS(responsive = false, htmlStyles = {}) {
        const {
            headerRowBold = true,
            headerColBold = false,
            alternatingRows = true,
            preset = 'default'
        } = htmlStyles;

        let css = '';

        // 根据预设选择基础样式
        switch (preset) {
            case 'minimal':
                css = this.getMinimalTableCSS();
                break;
            case 'professional':
                css = this.getProfessionalTableCSS();
                break;
            case 'modern':
                css = this.getModernTableCSS();
                break;
            default:
                css = this.getDefaultTableCSS(false); // 使用原有默认样式作为基础
                break;
        }

        // 添加增强样式
        let enhancedStyles = '';

        // 标题行加粗
        if (headerRowBold) {
            enhancedStyles += `
thead th {
    font-weight: bold !important;
}
`;
        }

        // 标题列加粗
        if (headerColBold) {
            enhancedStyles += `
.header-col {
    font-weight: bold !important;
}
`;
        }

        // 阴阳行背景
        if (alternatingRows) {
            enhancedStyles += `
.even-row {
    background-color: #ffffff;
}

.odd-row {
    background-color: #f8f9fa;
}

.even-row:hover,
.odd-row:hover {
    background-color: #e3f2fd;
}
`;
        }

        css += enhancedStyles;

        // 响应式样式
        if (responsive) {
            css += this.getResponsiveTableCSS();
        }

        return css;
    }

    /**
     * 获取默认的表格CSS样式
     */
    getDefaultTableCSS(responsive = false) {
        let css = `
table {
    border-collapse: collapse;
    width: 100%;
    margin: 20px 0;
    font-family: Arial, sans-serif;
}

th, td {
    border: 1px solid #ddd;
    padding: 8px 12px;
    text-align: left;
    vertical-align: top;
}

th {
    background-color: #f2f2f2;
    font-weight: bold;
}

tr:nth-child(even) {
    background-color: #f9f9f9;
}

tr:hover {
    background-color: #f5f5f5;
}
`;

        return css;
    }

    /**
     * 获取简约风格CSS
     */
    getMinimalTableCSS() {
        return `
table {
    border-collapse: collapse;
    width: 100%;
    margin: 20px 0;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

th, td {
    border: none;
    border-bottom: 1px solid #e0e0e0;
    padding: 12px 8px;
    text-align: left;
    vertical-align: top;
}

th {
    border-bottom: 2px solid #333;
    font-weight: 500;
}
`;
    }

    /**
     * 获取专业风格CSS
     */
    getProfessionalTableCSS() {
        return `
table {
    border-collapse: collapse;
    width: 100%;
    margin: 20px 0;
    font-family: 'Times New Roman', serif;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}

th, td {
    border: 1px solid #ccc;
    padding: 10px 15px;
    text-align: left;
    vertical-align: top;
}

th {
    background-color: #2c3e50;
    color: white;
    font-weight: bold;
    text-transform: uppercase;
    font-size: 12px;
    letter-spacing: 1px;
}

tr:hover {
    background-color: #f8f9fa;
}
`;
    }

    /**
     * 获取现代风格CSS
     */
    getModernTableCSS() {
        return `
table {
    border-collapse: collapse;
    width: 100%;
    margin: 20px 0;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 4px 20px rgba(0,0,0,0.1);
}

th, td {
    border: none;
    padding: 15px 20px;
    text-align: left;
    vertical-align: top;
}

th {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    font-weight: 600;
    text-transform: uppercase;
    font-size: 13px;
    letter-spacing: 0.5px;
}

td {
    border-bottom: 1px solid #f0f0f0;
}

tr:hover td {
    background-color: #f8f9ff;
}
`;
    }

    /**
     * 获取响应式CSS
     */
    getResponsiveTableCSS() {
        return `
@media screen and (max-width: 768px) {
    .responsive-table {
        border: 0;
    }
    
    .responsive-table thead {
        display: none;
    }
    
    .responsive-table tr {
        border: 1px solid #ccc;
        display: block;
        margin-bottom: 10px;
        border-radius: 4px;
    }
    
    .responsive-table td {
        border: none;
        display: block;
        text-align: right;
        padding-left: 50%;
        position: relative;
        padding-top: 8px;
        padding-bottom: 8px;
    }
    
    .responsive-table td:before {
        content: attr(data-label) ": ";
        position: absolute;
        left: 6px;
        width: 45%;
        text-align: left;
        font-weight: bold;
    }
}
`;
    }

    /**
     * 转换驼峰命名为kebab命名
     */
    camelToKebab(str) {
        return str.replace(/([a-z0-9]|(?=[A-Z]))([A-Z])/g, '$1-$2').toLowerCase();
    }

    /**
     * HTML转义
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * Markdown转义
     */
    escapeMarkdown(text) {
        return text.replace(/[|\\]/g, '\\$&');
    }

    /**
     * Textile转义
     */
    escapeTextile(text) {
        return text.replace(/[|]/g, '&#124;');
    }

    /**
     * 获取文件扩展名
     */
    getFileExtension(format) {
        const extensions = {
            'html': '.html',
            'markdown': '.md',
            'textile': '.textile',
            'rst': '.rst'
        };
        return extensions[format] || '.txt';
    }

    /**
     * 获取MIME类型
     */
    getMimeType(format) {
        const mimeTypes = {
            'html': 'text/html',
            'markdown': 'text/markdown',
            'textile': 'text/x-textile',
            'rst': 'text/x-rst'
        };
        return mimeTypes[format] || 'text/plain';
    }
} 