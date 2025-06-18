class TableParser {
    constructor() {
        this.tableData = null;
        this.rawData = null;
    }

    /**
     * 从剪贴板读取数据并解析
     */
    async parseFromClipboard() {
        try {
            const clipboardItems = await navigator.clipboard.read();
            
            for (const clipboardItem of clipboardItems) {
                // 尝试读取HTML格式（优先，包含格式信息）
                if (clipboardItem.types.includes('text/html')) {
                    const htmlBlob = await clipboardItem.getType('text/html');
                    const htmlText = await htmlBlob.text();
                    this.rawData = { type: 'html', content: htmlText };
                    return this.parseHtmlTable(htmlText);
                }
                
                // 回退到纯文本格式
                if (clipboardItem.types.includes('text/plain')) {
                    const textBlob = await clipboardItem.getType('text/plain');
                    const textContent = await textBlob.text();
                    this.rawData = { type: 'text', content: textContent };
                    return this.parseTextTable(textContent);
                }
            }
            
            throw new Error('剪贴板中没有找到表格数据');
        } catch (error) {
            console.error('读取剪贴板失败:', error);
            
            // 如果剪贴板API失败，尝试使用传统方法
            try {
                const textContent = await navigator.clipboard.readText();
                this.rawData = { type: 'text', content: textContent };
                return this.parseTextTable(textContent);
            } catch (fallbackError) {
                throw new Error('无法访问剪贴板，请确保浏览器支持并已授权剪贴板访问');
            }
        }
    }

    /**
     * 解析HTML表格
     */
    parseHtmlTable(htmlContent) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlContent, 'text/html');
        
        // 查找表格元素
        let table = doc.querySelector('table');
        
        if (!table) {
            // 如果没有table标签，尝试从其他元素中提取表格数据
            const rows = this.extractRowsFromHtml(doc);
            if (rows.length > 0) {
                return this.buildTableData(rows);
            }
            throw new Error('HTML中没有找到表格数据');
        }

        return this.parseTableElement(table);
    }

    /**
     * 解析table元素
     */
    parseTableElement(table) {
        const rows = [];
        const tableRows = table.querySelectorAll('tr');

        for (let rowIndex = 0; rowIndex < tableRows.length; rowIndex++) {
            const row = tableRows[rowIndex];
            const cells = [];
            const cellElements = row.querySelectorAll('td, th');

            for (let cellIndex = 0; cellIndex < cellElements.length; cellIndex++) {
                const cell = cellElements[cellIndex];
                const cellData = {
                    content: this.cleanCellContent(cell.textContent || cell.innerText || ''),
                    type: cell.tagName.toLowerCase(),
                    colspan: parseInt(cell.getAttribute('colspan')) || 1,
                    rowspan: parseInt(cell.getAttribute('rowspan')) || 1,
                    styles: this.extractCellStyles(cell)
                };

                cells.push(cellData);
            }

            if (cells.length > 0) {
                rows.push({
                    cells: cells,
                    type: this.determineRowType(cells)
                });
            }
        }

        this.tableData = this.buildTableData(rows);
        return this.tableData;
    }

    /**
     * 从HTML中提取行数据（针对非标准表格格式）
     */
    extractRowsFromHtml(doc) {
        const rows = [];
        
        // 尝试从各种可能的结构中提取数据
        const possibleRows = doc.querySelectorAll('tr, .row, [style*="display"][style*="table-row"]');
        
        for (const row of possibleRows) {
            const cells = this.extractCellsFromElement(row);
            if (cells.length > 0) {
                rows.push({
                    cells: cells,
                    type: 'data'
                });
            }
        }

        return rows;
    }

    /**
     * 从元素中提取单元格数据
     */
    extractCellsFromElement(element) {
        const cells = [];
        const cellElements = element.querySelectorAll('td, th, .cell, [style*="display"][style*="table-cell"]');

        for (const cell of cellElements) {
            const cellData = {
                content: this.cleanCellContent(cell.textContent || cell.innerText || ''),
                type: cell.tagName.toLowerCase() === 'th' ? 'th' : 'td',
                colspan: parseInt(cell.getAttribute('colspan')) || 1,
                rowspan: parseInt(cell.getAttribute('rowspan')) || 1,
                styles: this.extractCellStyles(cell)
            };
            cells.push(cellData);
        }

        return cells;
    }

    /**
     * 解析纯文本表格
     */
    parseTextTable(textContent) {
        const lines = textContent.trim().split(/\r?\n/);
        const rows = [];

        for (const line of lines) {
            if (line.trim() === '') continue;

            // 检测分隔符
            const separators = ['\t', ',', '|', ';'];
            let bestSeparator = '\t';
            let maxCells = 0;

            for (const sep of separators) {
                const cellCount = line.split(sep).length;
                if (cellCount > maxCells) {
                    maxCells = cellCount;
                    bestSeparator = sep;
                }
            }

            const cellContents = line.split(bestSeparator);
            const cells = cellContents.map(content => ({
                content: this.cleanCellContent(content),
                type: 'td',
                colspan: 1,
                rowspan: 1,
                styles: {}
            }));

            if (cells.length > 0) {
                rows.push({
                    cells: cells,
                    type: rows.length === 0 ? 'header' : 'data'
                });
            }
        }

        this.tableData = this.buildTableData(rows);
        return this.tableData;
    }

    /**
     * 构建标准化的表格数据结构
     */
    buildTableData(rows) {
        if (rows.length === 0) {
            return { rows: [], maxColumns: 0, hasHeader: false };
        }

        // 计算最大列数
        let maxColumns = 0;
        rows.forEach(row => {
            let columnCount = 0;
            row.cells.forEach(cell => {
                columnCount += cell.colspan;
            });
            maxColumns = Math.max(maxColumns, columnCount);
        });

        // 处理合并单元格的占位
        const processedRows = this.processMergedCells(rows, maxColumns);

        return {
            rows: processedRows,
            maxColumns: maxColumns,
            hasHeader: this.detectHeaderRow(processedRows)
        };
    }

    /**
     * 处理合并单元格
     */
    processMergedCells(rows, maxColumns) {
        const grid = [];
        
        // 初始化网格
        for (let i = 0; i < rows.length; i++) {
            grid[i] = new Array(maxColumns).fill(null);
        }

        // 填充网格
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            let colIndex = 0;

            for (const cell of row.cells) {
                // 找到下一个空白位置
                while (colIndex < maxColumns && grid[rowIndex][colIndex] !== null) {
                    colIndex++;
                }

                if (colIndex >= maxColumns) break;

                // 填充单元格及其跨越的区域
                for (let r = 0; r < cell.rowspan && (rowIndex + r) < rows.length; r++) {
                    for (let c = 0; c < cell.colspan && (colIndex + c) < maxColumns; c++) {
                        if (r === 0 && c === 0) {
                            grid[rowIndex + r][colIndex + c] = cell;
                        } else {
                            grid[rowIndex + r][colIndex + c] = { 
                                isMerged: true, 
                                parentCell: cell,
                                content: '',
                                type: cell.type,
                                colspan: 1,
                                rowspan: 1,
                                styles: {}
                            };
                        }
                    }
                }

                colIndex += cell.colspan;
            }
        }

        // 转换回行格式
        return grid.map((gridRow, rowIndex) => ({
            cells: gridRow,
            type: rows[rowIndex] ? rows[rowIndex].type : 'data'
        }));
    }

    /**
     * 检测是否有表头行
     */
    detectHeaderRow(rows) {
        if (rows.length === 0) return false;
        
        const firstRow = rows[0];
        const headerIndicators = firstRow.cells.filter(cell => 
            cell && (cell.type === 'th' || 
            (cell.styles && (cell.styles.fontWeight === 'bold' || 
                           cell.styles.backgroundColor)))
        );

        return headerIndicators.length > firstRow.cells.length * 0.5;
    }

    /**
     * 确定行类型
     */
    determineRowType(cells) {
        const headerCells = cells.filter(cell => cell.type === 'th').length;
        return headerCells > cells.length * 0.5 ? 'header' : 'data';
    }

    /**
     * 清理单元格内容
     */
    cleanCellContent(content) {
        return content
            .replace(/^\s+|\s+$/g, '') // 去除首尾空白
            .replace(/\s+/g, ' ') // 合并多个空格
            .replace(/[\r\n]+/g, ' '); // 替换换行为空格
    }

    /**
     * 提取单元格样式
     */
    extractCellStyles(cell) {
        const styles = {};
        const computedStyle = window.getComputedStyle(cell);

        // 提取重要样式
        const importantStyles = [
            'backgroundColor', 'color', 'fontWeight', 'fontStyle', 
            'textAlign', 'verticalAlign', 'fontSize', 'borderColor'
        ];

        importantStyles.forEach(style => {
            const value = computedStyle.getPropertyValue(this.camelToKebab(style));
            if (value && value !== 'initial' && value !== 'normal') {
                styles[style] = value;
            }
        });

        return styles;
    }

    /**
     * 驼峰命名转kebab命名
     */
    camelToKebab(str) {
        return str.replace(/([a-z0-9]|(?=[A-Z]))([A-Z])/g, '$1-$2').toLowerCase();
    }

    /**
     * 获取解析后的表格数据
     */
    getTableData() {
        return this.tableData;
    }

    /**
     * 获取原始数据
     */
    getRawData() {
        return this.rawData;
    }

    /**
     * 生成HTML表格用于预览
     */
    generatePreviewHtml(options = {}) {
        if (!this.tableData || !this.tableData.rows) {
            return '<div class="empty-state"><p>暂无表格数据</p></div>';
        }

        const { preserveMerged = true, includeStyles = true } = options;
        
        let html = '<table>';
        
        for (let rowIndex = 0; rowIndex < this.tableData.rows.length; rowIndex++) {
            const row = this.tableData.rows[rowIndex];
            html += '<tr>';
            
            for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
                const cell = row.cells[colIndex];
                
                if (!cell || cell.isMerged) continue;
                
                const tag = cell.type === 'th' ? 'th' : 'td';
                let cellHtml = `<${tag}`;
                
                // 添加合并属性
                if (preserveMerged) {
                    if (cell.colspan > 1) cellHtml += ` colspan="${cell.colspan}"`;
                    if (cell.rowspan > 1) cellHtml += ` rowspan="${cell.rowspan}"`;
                }
                
                // 添加样式
                if (includeStyles && cell.styles && Object.keys(cell.styles).length > 0) {
                    const styleStr = Object.entries(cell.styles)
                        .map(([key, value]) => `${this.camelToKebab(key)}: ${value}`)
                        .join('; ');
                    cellHtml += ` style="${styleStr}"`;
                }
                
                cellHtml += `>${this.escapeHtml(cell.content)}</${tag}>`;
                html += cellHtml;
            }
            
            html += '</tr>';
        }
        
        html += '</table>';
        return html;
    }

    /**
     * HTML转义
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
} 