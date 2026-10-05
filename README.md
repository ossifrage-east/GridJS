# SilkSpaces 藕丝空间

> 会道者，一缕藕丝牵大象。寡道者，千钧铁棒打苍蝇。
> —— The wise, with a lotus filament, can lead an elephant. The unwise, with an iron club, can but beat flies.

**SilkSpaces（藕丝空间）** —— GridJS
 是一个轻量级的 Web 电子表格组件，使用原生 JavaScript/TypeScript + HTML5 Canvas 实现，尽量不依赖第三方库，在浏览器中提供类似 Excel 的表格编辑体验。

![License](https://img.shields.io/badge/license-MIT-blue)
![TypeScript](https://img.shields.io/badge/language-TypeScript-blue)
![Canvas](https://img.shields.io/badge/render-HTML5_Canvas-orange)
![IndexedDB](https://img.shields.io/badge/storage-IndexedDB-green)

## 特性

- **Canvas 高性能渲染**：基于 Canvas 2D 的表格绘制，原生实现，几乎零第三方运行时依赖
- **单元格编辑**：光标控制、文本选择、输入法组合输入、多行文本支持
- **字符级格式**：单元格内每个字符可独立设置字体、字号、颜色、加粗、斜体、下划线、删除线、上下标、字间距、行间距
- **行列操作**：插入、删除、隐藏/取消隐藏，拖拽调整行高列宽，合并单元格
- **选区与剪贴板**：框选、全选、复制/剪切/粘贴
- **撤销/重做**：完整的文档级与编辑内撤销/重做支持
- **排序与筛选**：混合类型智能排序（时间 → 数字 → 英文 → 中文），数据筛选
- **打印预览**：纸张大小/方向、页边距、缩放、页面范围、双面打印等完整打印设置
- **缩放**：表格与打印预览双向同步缩放
- **右键菜单**：功能完整的上下文菜单
- **本地持久化**：数据自动保存至浏览器 IndexedDB
- **字体管理**：自定义字体检测与加载（内置 HarmonyOS Sans SC）
- **事件驱动**：EventEmitter 数据变更通知机制，模块松耦合
- **宏编程**：支持使用 Python 进行宏编程（规划中）

## 快速开始

### 环境要求

- Node.js（建议 16+）
- npm

### 安装与运行

```bash
# 安装依赖
npm install

# 启动开发服务器（终端会输出访问地址）
npm run dev

# 生产构建（输出至 dist/ 目录）
npm run build
```

## 项目结构

```
silkspaces/
├── src/
│   ├── assets/                    # 静态资源（字体、图标、光标等）
│   ├── core/
│   │   ├── common/                # 通用 UI 组件（菜单、滚动条、分割面板等）
│   │   ├── dataArchitecture/      # 数据架构层（单元格、字符、数据集合、IndexedDB）
│   │   ├── sheets/                # Canvas 渲染层（主表格/行头/列头/编辑画布）
│   │   ├── style/                 # SCSS 样式
│   │   ├── toolbar/               # 工具栏（快速访问栏、开始标签页各分区）
│   │   ├── app.ts                 # 应用入口类
│   │   └── constant.ts            # 常量与事件定义
│   ├── utils/                     # 工具模块（撤销/重做、剪贴板、字体加载、事件等）
│   └── index.ts                   # 启动入口
├── CodeWiki.md                    # 架构详解文档
├── LICENSE.md                     # MIT 协议
└── webpack.config.js              # Webpack 构建配置
```

## 技术栈

| 分类 | 技术 |
|------|------|
| 语言 | TypeScript |
| 渲染 | HTML5 Canvas 2D |
| 数据存储 | IndexedDB（原生 API 封装） |
| 构建 | Webpack 5 + ts-loader |
| 样式 | SCSS + PostCSS + Autoprefixer |

## 文档

详细的架构设计、模块说明与数据流分析请查阅 [CodeWiki.md](./CodeWiki.md)。

## 作者

**东方鹗 (eastossifrage)**

- B站：https://space.bilibili.com/194359739
- 知乎：https://www.zhihu.com/people/eastossifrage
- CSDN：https://blog.csdn.net/os373

## 开源协议

本项目基于 [MIT](./LICENSE.md) 协议开源。
