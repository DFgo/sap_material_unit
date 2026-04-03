# SAP 物料单位数据工具

用于合并 SAP 物料主数据（MARA）和物料单位转换数据（MARM），生成带双层表头的 Excel 文件，并支持邮件导出。

## 功能特性

- 上传 SAP SE16N 导出的 Excel 文件（material_data / material_unit）
- 自动合并物料主数据与单位转换数据
- 支持在线编辑、搜索筛选、分页浏览
- 导出带双层合并表头的格式化 Excel
- 生成 .eml 邮件文件（附件为 Excel）

## 技术栈

- Vue 3 + Composition API
- Vite
- Element Plus
- ExcelJS - Excel 处理（支持单元格样式）

## 本地运行

```sh
npm install
npm run dev
npm run build
```

---

## 数据来源（SAP SE16N 导出）

### 导出步骤

1. **登录 SAP GUI**
2. **事务码 SE16N**
3. **输入表名**，回车
4. **筛选数据**（可选）：在屏幕上设置过滤条件
5. **导出**：菜单 `System → List → Save → Local File → Spreadsheet`（或直接按 F9）

> 导出格式选择 **.xlsx** 或 **.xls** 均可

---

### material_data.xlsx — MARA 表

| 字段名 | SAP 字段 | 说明 |
|--------|----------|------|
| 物料 | MATNR | 物料编号 |
| 物料描述 | MAKTX | 物料描述（来自 MAKT 表，导出时合并） |
| 基本计量单位 | MEINS | 基本计量单位 |

**列索引对应关系**（SE16N 导出后）：

| Excel 列 | 字段 |
|----------|------|
| 第 1 列（A） | 物料 |
| 倒数第 1 列 | 物料描述 |
| 第 13 列（M） | 基本计量单位 |

---

### material_unit.xlsx — MARM 表

| 字段名 | SAP 字段 | 说明 |
|--------|----------|------|
| 物料 | MATNR | 物料编号 |
| 可选计量单位 | MSEHI | 可选计量单位 |
| 计数器 | UMREZ | 计数器（分子） |
| 分母 | UMREN | 分母 |

**列索引对应关系**（SE16N 导出后）：

| Excel 列 | 字段 |
|----------|------|
| 第 1 列（A） | 物料 |
| 第 2 列（B） | 可选计量单位 |
| 第 3 列（C） | 计数器 |
| 第 4 列（D） | 分母 |

---

### SE16N 导出示意

```
SE16N → 输入表名 → 回车 → F9导出 → 保存为本地文件
        ↓
   material_data.xlsx  ← MARA 表 + MAKT 物料描述
   material_unit.xlsx  ← MARM 表
```

> **注意**：请确保 MARA 和 MARM 导出的数据范围一致（相同的物料编号集合），否则部分物料可能无法匹配。

---

## 使用流程

1. **上传 material_data.xlsx**（MARA 物料主数据）
2. **上传 material_unit.xlsx**（MARM 单位转换数据）
3. **点击"开始计算"** — 系统自动合并数据
4. **查看合并结果** — 在表格中编辑、搜索
5. **导出 Excel** 或 **导出邮件**

---

## 输出文件格式

### 合并后表格列顺序（全局统一）

```
序号 | 物料 | 物料描述 | 基本计量单位 | 分母1 | 可选单位1 | = | 计数器1 | 基本计量单位1 | 分母2 | 可选单位2 | = | 计数器2 | 基本计量单位2 | ...
```

### Excel 导出（双层表头）

- 第 1 行：分组大标题（物料 / 物料描述 / 基本计量单位 / 单位转换1 / 单位转换2 / ...）
- 第 2 行：具体列名（分母 / 可选单位 / = / 计数器 / 基本计量单位）
- 前 3 列（A/B/C）：垂直合并居中
- 单位转换列组：每组 5 列水平合并居中
- 全表格：居中对齐 + 全边框

### 邮件导出（.eml）

- HTML 预览：显示前 20 行 + "此为数据预览，详见附件" 备注
- 附件：完整数据的 Excel 文件（格式与导出 Excel 一致）

---

## 项目结构

```
sap_material_unit/
├── src/
│   ├── main.js                  # 应用入口
│   ├── App.vue                   # 根组件
│   ├── router/index.js           # 路由配置
│   ├── utils/
│   │   └── excelProcessor.js     # Excel 读取/合并/导出核心逻辑
│   └── views/
│       └── MaterialMergeView.vue # 主页面
├── test/
│   └── data/
│       ├── se16n-mara.xlsx       # SAP SE16N MARA 导出样例
│       └── se16n-marm.xlsx        # SAP SE16N MARM 导出样例
└── docs/                         # 内部设计文档（不上传）
```

## License

MIT
