# 智能战友PPT生成器 v3.1

> 不是工具，是你的战略伙伴。预判你的预判，分析你的受众，优化你的效果。

## 三重AI协作架构

```
┌─────────────────────────────────────────────────────────┐
│  Gemini 3 Pro Preview                                   │
│  负责: PPT框架、结构、大纲、逻辑流程                     │
│  • 麦肯锡/BCG级别的结构化思维                           │
│  • 金字塔原理、MECE、SCQA框架                           │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│  Claude Opus 4.5                                        │
│  负责: 内容填充、文案润色、质量检查                      │
│  • 严格按照Gemini框架填充内容                           │
│  • 专业商务文案                                         │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│  Gemini 2.0 Flash                                       │
│  负责: 8K超高清渲染 (7680x4320)                         │
│  • 中文/英文/数字清晰呈现                               │
└─────────────────────────────────────────────────────────┘
```

## 快速开始

### 方式1: 使用快捷脚本

```bash
cd skills/ppt-generator
./generate.sh "2025年Q1业绩分析报告"
```

### 方式2: 直接运行Python

```bash
cd skills/ppt-generator/python
pip install -r requirements.txt
python intelligent_ppt_strategist.py -t "报告主题"
```

### 方式3: 完整参数

```bash
python intelligent_ppt_strategist.py \
  -t "2025年Q1业绩分析" \
  -c "销售额6.55亿，增长546%" \
  -a "c_suite" \
  -g "report"
```

## 参数说明

| 参数 | 说明 | 可选值 |
|------|------|--------|
| `-t, --topic` | 报告主题 | 任意文本 |
| `-c, --content` | 内容描述 | 任意文本 |
| `-a, --audience` | 受众类型 | 见下表 |
| `-g, --goal` | 演示目标 | 见下表 |
| `-d, --from-dashboard` | 使用仪表盘数据 | - |
| `--no-analysis` | 跳过智能分析 | - |
| `--no-prediction` | 跳过效果预判 | - |

### 受众类型

| 值 | 说明 |
|----|------|
| `c_suite` | CEO/CFO/CTO等高管 |
| `board` | 董事会 |
| `investors` | 投资人/VC/PE |
| `clients` | 客户/甲方 |
| `internal` | 内部团队 |
| `government` | 政府/监管机构 |
| `sales` | 销售场景 |
| `training` | 培训场景 |

### 演示目标

| 值 | 说明 |
|----|------|
| `report` | 汇报工作 |
| `pitch` | 融资/商业路演 |
| `persuade` | 说服/获取支持 |
| `inform` | 传达信息 |
| `teach` | 教学/培训 |

## 7阶段生成流程

1. **Phase 1: Claude战略分析** - 受众分析、目标识别、关键词提取
2. **Phase 2: Gemini框架设计** - 麦肯锡级PPT结构设计 (核心!)
3. **Phase 3: 智能风格选择** - 根据受众+目标自动匹配风格
4. **Phase 4: Claude内容填充** - 按框架填充专业文案
5. **Phase 5: Gemini 8K渲染** - 7680x4320超高清图像
6. **Phase 6: PPT组装** - 生成可编辑PPTX文件
7. **Phase 7: 效果预判** - 预测受众反应

## 与Manus对比

| 功能 | Manus | 智能战友 |
|------|-------|----------|
| 最大分辨率 | 4K (4096px) | **8K (7680px)** |
| 页数限制 | 12页 | **无限制** |
| 订阅费用 | $40/月 | **$0** |
| API等级 | 入门级 | **顶配Opus 4.5** |
| 数据时效 | 可能2024年 | **强制2025年** |
| 实时搜索 | ❌ | **Perplexity** |
| 网页抓取 | ❌ | **Firecrawl** |
| 受众分析 | ❌ | **智能预判** |
| 效果预判 | ❌ | **预测反馈** |
| 框架设计 | 通用模板 | **Gemini定制** |

## 环境配置

在项目根目录创建 `.env` 文件:

```env
# Gemini API (框架 + 图像)
GEMINI_API_KEY=your_gemini_api_key

# Claude API (内容填充)
CLAUDE_API_KEY=your_claude_api_key

# Perplexity API (实时搜索) - 可选
PERPLEXITY_API_KEY=your_perplexity_key

# Firecrawl API (网页抓取) - 可选
FIRECRAWL_API_KEY=your_firecrawl_key
```

## 输出

生成的PPT保存在 `./output/` 目录:

```
{主题}_智能_{页数}页_{时间戳}.pptx
```

## 框架模板

系统内置6种商业框架模板:

| 框架 | 适用场景 | 原则 |
|------|----------|------|
| executive_report | 高管汇报 | 金字塔原理 |
| investor_pitch | 投资路演 | 故事线叙事 |
| sales_proposal | 销售提案 | SPIN销售 |
| strategy_review | 战略回顾 | SWOT分析 |
| project_update | 项目汇报 | RAG状态 |
| training | 培训教学 | 布鲁姆层次 |

## 目录结构

```
skills/ppt-generator/
├── generate.sh           # 快捷启动脚本
├── README.md            # 本文件
├── python/
│   ├── intelligent_ppt_strategist.py  # 主程序 v3.1
│   ├── enterprise_ppt_generator.py    # 企业版(含Perplexity)
│   ├── manus_clone_ppt.py             # Manus克隆版
│   ├── nanobanana_ppt.py              # 基础版
│   └── requirements.txt               # Python依赖
└── index.js             # Node.js版本
```

## Claude Max导入

将 `.claude/` 目录复制到你的项目中:

```bash
cp -r .claude/ ~/your-project/
```

然后在Claude Code中使用:
```
/ppt 2025年Q1业绩分析
```
