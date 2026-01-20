# PPT Generator Skill

使用 **Nano Banana Pro** (Gemini 3 Pro) 生成高质量PPT演示文稿的技能模块。

## 特性

- 调用 Nano Banana Pro API 生成精美幻灯片图像
- 支持中文文本渲染，清晰可读
- 多种预设风格：专业商务、极简白色、活力渐变
- 支持从仪表盘数据自动生成报告
- 提供 CLI 工具和 REST API 接口

## 快速开始

### 1. 安装依赖

```bash
cd gree-dashboard
npm install
```

### 2. 配置 API 密钥

复制环境变量示例文件并填入您的 Gemini API 密钥：

```bash
cp .env.example .env
```

编辑 `.env` 文件：

```env
GEMINI_API_KEY=your_gemini_api_key_here
```

> 获取 API 密钥: https://aistudio.google.com/app/apikey

### 3. 生成 PPT

#### CLI 方式

```bash
# 从仪表盘数据生成
npm run generate-ppt -- --from-dashboard

# 从 JSON 文件生成
npm run generate-ppt -- -t "我的演示" -s skills/ppt-generator/examples/slides.json

# 查看帮助
npm run generate-ppt -- --help
```

#### API 方式

启动服务器：

```bash
npm start
```

调用 API：

```bash
# 生成 PPT
curl -X POST http://localhost:3000/api/generate \
  -H "Content-Type: application/json" \
  -d '{
    "title": "测试演示",
    "slides": [
      {"title": "封面", "subtitle": "副标题"},
      {"title": "内容页", "bullets": ["要点1", "要点2"]}
    ],
    "style": "professional"
  }'
```

## API 接口

| 端点 | 方法 | 描述 |
|------|------|------|
| `/api/health` | GET | 健康检查 |
| `/api/styles` | GET | 获取可用风格 |
| `/api/generate` | POST | 生成 PPT |
| `/api/generate/dashboard` | POST | 从仪表盘数据生成 |
| `/api/generate/slide-image` | POST | 生成单张幻灯片图像 |
| `/api/download/:fileName` | GET | 下载 PPT 文件 |
| `/api/files` | GET | 列出已生成文件 |

## 幻灯片格式

```json
[
  {
    "title": "标题",
    "subtitle": "副标题"
  },
  {
    "title": "内容页标题",
    "content": "正文内容",
    "bullets": ["要点1", "要点2", "要点3"]
  },
  "简单文本幻灯片"
]
```

## 可用风格

- **professional**: 专业商务 - 深蓝色背景，青色和白色文字
- **minimal**: 极简白色 - 白色背景，黑色文字
- **vibrant**: 活力渐变 - 深色背景配亮色文字

## 编程接口

```javascript
import { PPTGenerator } from './skills/ppt-generator/index.js';

const generator = new PPTGenerator(process.env.GEMINI_API_KEY);

// 生成 PPT
const outputPath = await generator.generate({
  title: '我的演示',
  slides: [
    { title: '封面', subtitle: '使用 Nano Banana Pro' },
    { title: '内容', bullets: ['要点1', '要点2'] }
  ],
  style: 'professional',
  useNanoBanana: true
});

// 从仪表盘数据生成
const reportPath = await generator.generateFromDashboard({
  title: '数据分析报告',
  kpis: [{ label: '销售额', value: '¥100万' }],
  charts: [{ title: '趋势图', highlights: ['增长10%'] }],
  insights: ['核心洞察1', '核心洞察2']
});
```

## 技术栈

- **Nano Banana Pro** (Gemini 3 Pro): Google 最新图像生成模型
- **pptxgenjs**: PowerPoint 文件生成
- **Express**: REST API 服务
- **Node.js**: 运行时环境

## 关于 Nano Banana Pro

Nano Banana Pro 是 Google DeepMind 推出的最新图像生成模型，基于 Gemini 3 Pro 架构。其特点：

- 能够准确渲染清晰、可读的文本
- 专门优化了演示文稿格式的生成
- 支持多语言文本渲染
- 生成的图像质量高，适合专业场景

更多信息：https://blog.google/technology/ai/nano-banana-pro/

## License

MIT
