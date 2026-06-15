# 公众号文章批量下载与术语表生成工具

一款基于 PyQt5 的桌面应用，支持批量下载微信公众号文章、智能提取术语、AI 自动翻译，并导出为 Excel/CSV/Markdown 等多种格式。

<p align="center">
  <strong>© 2026 June | Lingoes Family | Vibe Coding</strong>
</p>

---

## 功能特性

### 文章下载
- 批量下载微信公众号文章，自动保存为 Markdown 格式
- 支持包含/排除关键词过滤
- 多线程并发下载，提升效率
- 支持从 cookies 文件加载登录状态

### 术语表生成
- **顺序提取模式**：按文章顺序提取中英文对照词对，适合词汇列表类文章
- **频率统计模式**：统计英文单词出现频率，适合普通英文文章
- **NLP 模式**：基于词性标注优先提取名词和形容词
- **AI 术语提取**：通过 DeepSeek API 智能理解文章并自动归纳术语

### 术语翻译
- 支持免费翻译引擎（无需配置即可使用）
- 支持 DeepSeek API 翻译（翻译质量更高，可批量翻译）
- 自动识别并输出英文-中文对照格式

### 多格式导出
- Excel (.xlsx) — 带表头样式、边框、列宽
- CSV (.csv) — 兼容 Excel 打开
- Markdown (.md) — 表格格式
- JSON (.json) — 结构化数据

---

## 快速开始

### 环境要求
- Python 3.8+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install PyQt5 requests beautifulsoup4 wxmp nltk openpyxl
```

或使用项目本地 libs 目录（已包含部分依赖）：

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
python wedown_api.py
```

---

## 使用说明

### 术语表生成

1. **输入文章内容**：在"文章输入"区域粘贴文章文本
2. **选择提取模式**：
   - 顺序提取模式：适合"1. 贵族aristocracy"这类词汇列表
   - 频率统计模式：适合普通英文文章
3. **调整参数**：设置最小词长度、最小出现次数、最大术语数量等
4. **点击提取**：点击"提取术语"按钮生成术语表
5. **AI 提取**：配置 DeepSeek API Key 后，点击"AI术语提取"按钮智能归纳
6. **翻译**：点击"识别并翻译术语"自动补全中文翻译
7. **保存**：点击"保存术语表"导出为 Excel/CSV/Markdown/JSON

### 文章下载

1. 在"文章下载"标签页输入公众号文章 URL（每行一个）
2. 配置微信公众平台 cookies（wxuin 和 pass_ticket）
3. 设置保存目录和并发数
4. 点击"开始下载"

---

## 项目结构

```
wedown/
── wedown_api.py          # 主程序（带术语翻译功能）
├── libs/                  # 本地依赖库
├── down_samples/          # 示例文件
└── README.md              # 说明文档
```

---

## 依赖说明

| 库 | 用途 | 是否必需 |
|---|---|---|
| PyQt5 | GUI 界面 | 必需 |
| requests | HTTP 请求 | 必需 |
| beautifulsoup4 | HTML 解析 | 必需 |
| wxmp | 微信公众号 API | 下载功能必需 |
| nltk | 自然语言处理 | NLP 模式必需 |
| openpyxl | Excel 导出 | Excel 导出必需 |

---

## 常见问题

**Q: 提示缺少某个库？**
A: 运行 `pip install <库名>` 安装对应依赖。

**Q: AI 术语提取失败？**
A: 请检查 DeepSeek API Key 是否正确，并确保网络连接正常。

**Q: 下载功能不可用？**
A: 请安装 wxmp 库：`pip install wxmp`，并正确配置 cookies。

---

## 许可证

本项目仅供学习交流使用。

---

## 致谢

© 2026 June | Lingoes Family | Vibe Coding
