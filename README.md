# SAP Material Unit Excel 处理工具

一个纯前端的 Vue 3 单页面应用，用于处理 SAP 物料单位转换数据的合并与导出。

## 功能特性

- 📤 **上传 Excel 文件** - 分别上传 material_data.xlsx 和 material_unit.xlsx
- 🔄 **自动合并数据** - 根据物料编号关联两个文件的数据
- ✏️ **在线编辑** - 支持在表格中直接修改数据
- 🔍 **搜索筛选** - 按物料编号或描述快速查找
- 📊 **导出 Excel** - 生成带合并标题的格式化 Excel 文件
- 📧 **邮件附件** - 生成 Base64 编码的邮件附件

## 技术栈

- Vue 3 + Composition API
- Vite
- Element Plus
- SheetJS (xlsx) - Excel 处理

## 本地运行

```sh
# 安装依赖
npm install

# 启动开发服务器
npm run dev

# 构建生产版本
npm run build
```

## 使用说明

### 输入文件格式

**material_data.xlsx**
| 列索引 | 字段 |
|--------|------|
| 0 | 物料 |
| 倒数第1列 | 物料描述 |
| 12 | 基本计量单位 |

**material_unit.xlsx**
| 列索引 | 字段 |
|--------|------|
| 0 | 物料 |
| 1 | 可选计量单位 |
| 2 | 计数器 |
| 3 | 分母 |

### 操作流程

1. 上传 `material_data.xlsx`
2. 上传 `material_unit.xlsx`
3. 数据自动合并后显示在表格中
4. 可直接编辑、搜索、添加/删除行
5. 点击"导出"下载 Excel 或"邮件"生成附件

## 项目结构

```
src/
├── main.js              # 应用入口
├── App.vue              # 根组件
├── router/index.js      # 路由配置
├── utils/
│   └── excelProcessor.js  # Excel 处理核心逻辑
└── views/
    └── MaterialMergeView.vue  # 主页面
```

## License

MIT
