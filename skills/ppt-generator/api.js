/**
 * PPT Generator API Server
 * 提供RESTful API接口用于PPT生成
 */

import express from 'express';
import { PPTGenerator, NanoBananaClient } from './index.js';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 加载环境变量
dotenv.config({ path: path.join(__dirname, '../../.env') });

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, '../../')));

// CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// 初始化生成器
let generator = null;

function getGenerator() {
  if (!generator) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error('GEMINI_API_KEY not configured');
    }
    generator = new PPTGenerator(apiKey, {
      outputDir: process.env.PPT_OUTPUT_DIR || path.join(__dirname, '../../output'),
      defaultStyle: process.env.PPT_DEFAULT_STYLE || 'professional'
    });
  }
  return generator;
}

/**
 * API健康检查
 */
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    service: 'PPT Generator Skill',
    version: '1.0.0',
    nanoBananaEnabled: !!process.env.GEMINI_API_KEY
  });
});

/**
 * 获取可用风格列表
 */
app.get('/api/styles', (req, res) => {
  res.json({
    styles: [
      {
        id: 'professional',
        name: '专业商务',
        description: '深蓝色背景，青色和白色文字，简洁现代'
      },
      {
        id: 'minimal',
        name: '极简白色',
        description: '白色背景，黑色文字，优雅简约'
      },
      {
        id: 'vibrant',
        name: '活力渐变',
        description: '深色背景配亮色文字，动感现代'
      }
    ]
  });
});

/**
 * 生成PPT
 * POST /api/generate
 *
 * Body:
 * {
 *   "title": "演示标题",
 *   "slides": [...],
 *   "style": "professional",
 *   "useNanoBanana": true
 * }
 */
app.post('/api/generate', async (req, res) => {
  try {
    const gen = getGenerator();
    const { title, slides, style, useNanoBanana = true, author } = req.body;

    if (!slides || !Array.isArray(slides) || slides.length === 0) {
      return res.status(400).json({
        error: 'Invalid request',
        message: 'slides array is required and must not be empty'
      });
    }

    const outputPath = await gen.generate({
      title: title || '演示文稿',
      slides,
      style: style || 'professional',
      useNanoBanana,
      author
    });

    // 读取生成的文件
    const fileBuffer = await fs.readFile(outputPath);
    const fileName = path.basename(outputPath);

    res.json({
      success: true,
      fileName,
      filePath: outputPath,
      downloadUrl: `/api/download/${fileName}`,
      slidesCount: slides.length
    });

  } catch (error) {
    console.error('Generate error:', error);
    res.status(500).json({
      error: 'Generation failed',
      message: error.message
    });
  }
});

/**
 * 从仪表盘数据生成PPT
 * POST /api/generate/dashboard
 */
app.post('/api/generate/dashboard', async (req, res) => {
  try {
    const gen = getGenerator();
    const dashboardData = req.body;

    if (!dashboardData.title) {
      dashboardData.title = '数据分析报告';
    }

    const outputPath = await gen.generateFromDashboard(dashboardData);
    const fileName = path.basename(outputPath);

    res.json({
      success: true,
      fileName,
      filePath: outputPath,
      downloadUrl: `/api/download/${fileName}`
    });

  } catch (error) {
    console.error('Dashboard generate error:', error);
    res.status(500).json({
      error: 'Generation failed',
      message: error.message
    });
  }
});

/**
 * 生成单张幻灯片图像 (Nano Banana Pro)
 * POST /api/generate/slide-image
 */
app.post('/api/generate/slide-image', async (req, res) => {
  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({
        error: 'API key not configured',
        message: 'GEMINI_API_KEY is required'
      });
    }

    const nanoBanana = new NanoBananaClient(apiKey);
    const { prompt, style = 'professional' } = req.body;

    if (!prompt) {
      return res.status(400).json({
        error: 'Invalid request',
        message: 'prompt is required'
      });
    }

    const imageBuffer = await nanoBanana.generateSlideImage(prompt, { style });

    res.json({
      success: true,
      image: imageBuffer.toString('base64'),
      mimeType: 'image/png'
    });

  } catch (error) {
    console.error('Slide image generation error:', error);
    res.status(500).json({
      error: 'Image generation failed',
      message: error.message
    });
  }
});

/**
 * 下载生成的PPT文件
 */
app.get('/api/download/:fileName', async (req, res) => {
  try {
    const outputDir = process.env.PPT_OUTPUT_DIR || path.join(__dirname, '../../output');
    const filePath = path.join(outputDir, req.params.fileName);

    // 安全检查：确保文件在output目录内
    const resolvedPath = path.resolve(filePath);
    const resolvedOutputDir = path.resolve(outputDir);

    if (!resolvedPath.startsWith(resolvedOutputDir)) {
      return res.status(403).json({ error: 'Access denied' });
    }

    // 检查文件是否存在
    await fs.access(filePath);

    res.download(filePath, req.params.fileName);

  } catch (error) {
    res.status(404).json({
      error: 'File not found',
      message: error.message
    });
  }
});

/**
 * 列出已生成的PPT文件
 */
app.get('/api/files', async (req, res) => {
  try {
    const outputDir = process.env.PPT_OUTPUT_DIR || path.join(__dirname, '../../output');

    await fs.mkdir(outputDir, { recursive: true });

    const files = await fs.readdir(outputDir);
    const pptFiles = files.filter(f => f.endsWith('.pptx'));

    const fileInfos = await Promise.all(
      pptFiles.map(async (fileName) => {
        const filePath = path.join(outputDir, fileName);
        const stat = await fs.stat(filePath);
        return {
          fileName,
          size: stat.size,
          createdAt: stat.birthtime,
          downloadUrl: `/api/download/${fileName}`
        };
      })
    );

    // 按创建时间排序（最新在前）
    fileInfos.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    res.json({ files: fileInfos });

  } catch (error) {
    res.status(500).json({
      error: 'Failed to list files',
      message: error.message
    });
  }
});

// 启动服务器
const PORT = process.env.PORT || 3000;

export function startServer(port = PORT) {
  return app.listen(port, () => {
    console.log(`
╔═══════════════════════════════════════════════════════════════╗
║          🍌 PPT Generator API Server                           ║
╚═══════════════════════════════════════════════════════════════╝

🚀 服务已启动: http://localhost:${port}

📡 API端点:
   GET  /api/health              - 健康检查
   GET  /api/styles              - 获取可用风格
   POST /api/generate            - 生成PPT
   POST /api/generate/dashboard  - 从仪表盘数据生成
   POST /api/generate/slide-image - 生成单张幻灯片图像
   GET  /api/download/:fileName  - 下载PPT文件
   GET  /api/files               - 列出已生成文件

🔑 Gemini API: ${process.env.GEMINI_API_KEY ? '已配置' : '❌ 未配置'}
    `);
  });
}

export default app;
