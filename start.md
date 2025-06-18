# 🚀 TableClip 快速启动指南

## 📁 项目结构
```
TableClip/
├── index.html              # 主应用页面
├── example.html            # 示例效果展示
├── README.md               # 详细说明文档
├── start.md               # 快速启动指南 (本文件)
├── css/
│   └── style.css          # 主样式文件
└── js/
    ├── tableParser.js     # 表格解析模块
    ├── formatExporters.js # 格式导出模块
    └── app.js            # 主应用逻辑
```

## ⚡ 立即使用

### 方法 1: 直接打开文件
1. 双击 `index.html` 文件
2. 浏览器会自动打开应用

### 方法 2: 本地服务器 (推荐)
```bash
# 使用 Python
python -m http.server 8000

# 使用 Node.js
npx serve .

# 使用 PHP
php -S localhost:8000
```

然后访问: http://localhost:8000

## 🎯 快速测试

### 1. 准备测试数据
在 Excel 中创建以下测试表格：
```
姓名    部门    年龄    城市
张三    销售    25     北京
李四    市场    30     上海
王五    技术    28     深圳
```

### 2. 测试步骤
1. 在 Excel 中选择表格区域 (包括表头)
2. 复制 (Ctrl+C)
3. 打开 TableClip 应用
4. 点击 "粘贴表格数据" 按钮
5. 选择输出格式 (HTML/Markdown/Textile/RST)
6. 查看预览和生成的代码

### 3. 功能验证
- ✅ 表格正确解析
- ✅ 预览显示正常
- ✅ 代码格式正确
- ✅ 复制/下载功能

## 🔧 开发调试

### 浏览器开发者工具
按 F12 打开开发者工具查看：
- Console: 查看错误日志
- Network: 检查资源加载
- Application: 查看剪贴板权限

### 常见问题排查
1. **粘贴按钮无响应**
   - 检查是否使用 HTTPS 或 localhost
   - 确认浏览器支持剪贴板 API
   - 检查剪贴板权限设置

2. **样式显示异常**
   - 确认 CSS 文件路径正确
   - 检查 Font Awesome 图标加载

3. **JavaScript 错误**
   - 确认所有 JS 文件正确加载
   - 检查浏览器兼容性

## 📝 自定义修改

### 修改主题色
编辑 `css/style.css`，修改 CSS 变量：
```css
body {
    background: linear-gradient(135deg, #你的颜色1, #你的颜色2);
}

.btn-primary {
    background: linear-gradient(135deg, #你的颜色1, #你的颜色2);
}
```

### 添加新的导出格式
1. 在 `formatExporters.js` 中添加新的导出方法
2. 在 `index.html` 中添加对应按钮
3. 在 `app.js` 中添加事件处理

## 🚢 部署上线

### GitHub Pages
1. 上传代码到 GitHub 仓库
2. 在设置中启用 GitHub Pages
3. 选择源分支 (通常是 main)

### 其他静态托管平台
- Netlify
- Vercel
- Surge.sh
- Firebase Hosting

## 📞 技术支持

如遇问题，请检查：
1. 浏览器控制台错误信息
2. 剪贴板数据格式
3. 网络连接状态

---

🎉 **恭喜！** 现在你可以开始使用 TableClip 了！ 