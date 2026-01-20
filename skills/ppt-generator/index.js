/**
 * PPT Generator Skill - 使用 Nano Banana Pro (Gemini 3 Pro) 生成PPT
 *
 * 功能:
 * - 调用 Nano Banana Pro 生成高质量幻灯片图像
 * - 支持文本内容自动排版
 * - 生成专业的商务演示文稿
 */

import { GoogleGenerativeAI } from '@google/generative-ai';
import PptxGenJS from 'pptxgenjs';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// PPT模板样式配置
const SLIDE_STYLES = {
  professional: {
    background: { color: '0F3460' },
    title: { color: '00D4FF', fontSize: 36, bold: true },
    subtitle: { color: 'FFFFFF', fontSize: 18 },
    body: { color: 'E0E0E0', fontSize: 14 },
    accent: '00FF88'
  },
  minimal: {
    background: { color: 'FFFFFF' },
    title: { color: '1A1A2E', fontSize: 36, bold: true },
    subtitle: { color: '666666', fontSize: 18 },
    body: { color: '333333', fontSize: 14 },
    accent: '0F3460'
  },
  vibrant: {
    background: { color: '1A1A2E' },
    title: { color: 'FFD93D', fontSize: 36, bold: true },
    subtitle: { color: 'FF6B6B', fontSize: 18 },
    body: { color: 'FFFFFF', fontSize: 14 },
    accent: '00D4FF'
  }
};

/**
 * Nano Banana Pro 客户端
 * 用于生成高质量的PPT幻灯片图像
 */
class NanoBananaClient {
  constructor(apiKey) {
    this.genAI = new GoogleGenerativeAI(apiKey);
    // 使用 Gemini 3 Pro (Nano Banana Pro) 模型
    this.model = this.genAI.getGenerativeModel({
      model: 'gemini-2.0-flash-exp-image-generation',
      generationConfig: {
        responseModalities: ['Text', 'Image']
      }
    });
  }

  /**
   * 生成幻灯片图像
   * @param {string} prompt - 幻灯片内容描述
   * @param {object} options - 生成选项
   * @returns {Promise<Buffer>} - 图像Buffer
   */
  async generateSlideImage(prompt, options = {}) {
    const {
      style = 'professional',
      aspectRatio = '16:9',
      language = 'zh-CN'
    } = options;

    const slidePrompt = this._buildSlidePrompt(prompt, style, aspectRatio, language);

    try {
      const result = await this.model.generateContent(slidePrompt);
      const response = result.response;

      // 提取生成的图像
      for (const part of response.candidates[0].content.parts) {
        if (part.inlineData) {
          const imageData = part.inlineData.data;
          return Buffer.from(imageData, 'base64');
        }
      }

      throw new Error('No image generated in response');
    } catch (error) {
      console.error('Nano Banana Pro API Error:', error.message);
      throw error;
    }
  }

  /**
   * 构建幻灯片生成提示词
   */
  _buildSlidePrompt(content, style, aspectRatio, language) {
    const styleGuides = {
      professional: '专业商务风格，深蓝色背景，青色和白色文字，简洁现代',
      minimal: '极简白色风格，大量留白，黑色文字，优雅简约',
      vibrant: '活力渐变风格，深色背景配亮色文字，动感现代'
    };

    return `Create a professional presentation slide image with the following specifications:

Content: ${content}

Style Requirements:
- Style: ${styleGuides[style] || styleGuides.professional}
- Aspect Ratio: ${aspectRatio}
- Language: ${language === 'zh-CN' ? 'Chinese (Simplified)' : 'English'}
- Text must be clear, readable, and professionally typeset
- Use high contrast for readability
- Include appropriate visual hierarchy
- Make it look like a polished corporate presentation slide

Important: Generate a complete slide image, not a template. The text should be rendered directly on the image with proper formatting.`;
  }

  /**
   * 批量生成多张幻灯片
   */
  async generateMultipleSlides(slides, options = {}) {
    const results = [];

    for (let i = 0; i < slides.length; i++) {
      console.log(`Generating slide ${i + 1}/${slides.length}...`);
      try {
        const imageBuffer = await this.generateSlideImage(slides[i], options);
        results.push({
          index: i,
          success: true,
          image: imageBuffer
        });
      } catch (error) {
        results.push({
          index: i,
          success: false,
          error: error.message
        });
      }

      // 添加延迟避免API限流
      if (i < slides.length - 1) {
        await this._delay(1000);
      }
    }

    return results;
  }

  _delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * PPT Generator 主类
 */
class PPTGenerator {
  constructor(apiKey, options = {}) {
    this.apiKey = apiKey;
    this.nanoBanana = new NanoBananaClient(apiKey);
    this.outputDir = options.outputDir || path.join(__dirname, '../../output');
    this.defaultStyle = options.defaultStyle || 'professional';
  }

  /**
   * 生成完整的PPT演示文稿
   * @param {object} config - PPT配置
   * @returns {Promise<string>} - 生成的文件路径
   */
  async generate(config) {
    const {
      title,
      author = 'Gree Dashboard',
      slides,
      style = this.defaultStyle,
      useNanoBanana = true,
      outputFileName
    } = config;

    console.log(`\n🎨 开始生成PPT: ${title}`);
    console.log(`📊 共 ${slides.length} 张幻灯片`);
    console.log(`🎯 风格: ${style}`);
    console.log(`🍌 Nano Banana Pro: ${useNanoBanana ? '启用' : '禁用'}\n`);

    // 确保输出目录存在
    await fs.mkdir(this.outputDir, { recursive: true });

    const pptx = new PptxGenJS();
    pptx.author = author;
    pptx.title = title;
    pptx.subject = 'Generated by Gree Dashboard PPT Skill';
    pptx.company = 'Gree Dashboard';

    // 设置幻灯片尺寸 (16:9)
    pptx.defineLayout({ name: 'CUSTOM_16x9', width: 10, height: 5.625 });
    pptx.layout = 'CUSTOM_16x9';

    const styleConfig = SLIDE_STYLES[style] || SLIDE_STYLES.professional;

    if (useNanoBanana) {
      // 使用 Nano Banana Pro 生成图像式幻灯片
      await this._generateImageSlides(pptx, slides, style);
    } else {
      // 使用传统模板方式生成幻灯片
      await this._generateTemplateSlides(pptx, slides, styleConfig);
    }

    // 生成文件名
    const fileName = outputFileName || `${title.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_')}_${Date.now()}.pptx`;
    const outputPath = path.join(this.outputDir, fileName);

    // 保存文件
    await pptx.writeFile({ fileName: outputPath });

    console.log(`\n✅ PPT生成成功!`);
    console.log(`📁 文件路径: ${outputPath}`);

    return outputPath;
  }

  /**
   * 使用 Nano Banana Pro 生成图像式幻灯片
   */
  async _generateImageSlides(pptx, slides, style) {
    // 先生成所有幻灯片图像
    const slideContents = slides.map(slide => {
      if (typeof slide === 'string') {
        return slide;
      }
      return `Title: ${slide.title}\n\nContent:\n${slide.content || slide.bullets?.join('\n') || ''}`;
    });

    const generatedImages = await this.nanoBanana.generateMultipleSlides(slideContents, { style });

    for (let i = 0; i < slides.length; i++) {
      const slide = pptx.addSlide();
      const generated = generatedImages[i];

      if (generated.success && generated.image) {
        // 将图像作为幻灯片背景
        const tempImagePath = path.join(this.outputDir, `temp_slide_${i}.png`);
        await fs.writeFile(tempImagePath, generated.image);

        slide.addImage({
          path: tempImagePath,
          x: 0,
          y: 0,
          w: '100%',
          h: '100%'
        });

        // 清理临时文件
        await fs.unlink(tempImagePath).catch(() => {});
      } else {
        // 如果图像生成失败，使用模板回退
        console.warn(`⚠️ Slide ${i + 1} image generation failed, using template fallback`);
        await this._addTemplateSlide(slide, slides[i], SLIDE_STYLES[style]);
      }
    }
  }

  /**
   * 使用传统模板方式生成幻灯片
   */
  async _generateTemplateSlides(pptx, slides, styleConfig) {
    for (const slideData of slides) {
      const slide = pptx.addSlide();
      await this._addTemplateSlide(slide, slideData, styleConfig);
    }
  }

  /**
   * 添加模板式幻灯片
   */
  async _addTemplateSlide(slide, slideData, styleConfig) {
    // 设置背景
    slide.background = styleConfig.background;

    if (typeof slideData === 'string') {
      // 简单文本幻灯片
      slide.addText(slideData, {
        x: 0.5,
        y: 2,
        w: 9,
        h: 1.5,
        fontSize: styleConfig.title.fontSize,
        color: styleConfig.title.color,
        bold: styleConfig.title.bold,
        align: 'center',
        valign: 'middle'
      });
    } else {
      // 结构化幻灯片
      const { title, subtitle, content, bullets, image, chart } = slideData;

      // 标题
      if (title) {
        slide.addText(title, {
          x: 0.5,
          y: 0.3,
          w: 9,
          h: 0.8,
          fontSize: styleConfig.title.fontSize,
          color: styleConfig.title.color,
          bold: styleConfig.title.bold,
          align: 'left'
        });
      }

      // 副标题
      if (subtitle) {
        slide.addText(subtitle, {
          x: 0.5,
          y: 1.0,
          w: 9,
          h: 0.5,
          fontSize: styleConfig.subtitle.fontSize,
          color: styleConfig.subtitle.color,
          align: 'left'
        });
      }

      // 正文内容
      if (content) {
        slide.addText(content, {
          x: 0.5,
          y: 1.6,
          w: 9,
          h: 3.5,
          fontSize: styleConfig.body.fontSize,
          color: styleConfig.body.color,
          align: 'left',
          valign: 'top'
        });
      }

      // 要点列表
      if (bullets && bullets.length > 0) {
        const bulletText = bullets.map(b => ({ text: b, options: { bullet: true } }));
        slide.addText(bulletText, {
          x: 0.5,
          y: 1.6,
          w: 9,
          h: 3.5,
          fontSize: styleConfig.body.fontSize,
          color: styleConfig.body.color,
          align: 'left',
          valign: 'top'
        });
      }

      // 图片
      if (image) {
        slide.addImage({
          path: image.path || image,
          x: image.x || 5,
          y: image.y || 1.5,
          w: image.w || 4,
          h: image.h || 3
        });
      }
    }
  }

  /**
   * 从仪表盘数据生成PPT
   * @param {object} dashboardData - 仪表盘数据
   * @returns {Promise<string>} - 生成的文件路径
   */
  async generateFromDashboard(dashboardData) {
    const {
      title = '数据分析报告',
      kpis = [],
      charts = [],
      insights = []
    } = dashboardData;

    const slides = [];

    // 封面页
    slides.push({
      title: title,
      subtitle: `生成时间: ${new Date().toLocaleDateString('zh-CN')}`,
      content: ''
    });

    // KPI概览页
    if (kpis.length > 0) {
      slides.push({
        title: '关键指标概览',
        bullets: kpis.map(kpi => `${kpi.label}: ${kpi.value}`)
      });
    }

    // 图表分析页
    for (const chart of charts) {
      slides.push({
        title: chart.title,
        content: chart.description || '',
        bullets: chart.highlights || []
      });
    }

    // 洞察总结页
    if (insights.length > 0) {
      slides.push({
        title: '关键洞察',
        bullets: insights
      });
    }

    // 结尾页
    slides.push({
      title: '感谢观看',
      subtitle: 'Thank You',
      content: '由 Gree Dashboard PPT Skill 生成\n使用 Nano Banana Pro (Gemini 3 Pro) 技术'
    });

    return this.generate({
      title,
      slides,
      useNanoBanana: true
    });
  }
}

export { PPTGenerator, NanoBananaClient, SLIDE_STYLES };
export default PPTGenerator;
