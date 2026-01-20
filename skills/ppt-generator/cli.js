#!/usr/bin/env node

/**
 * PPT Generator CLI
 * 命令行工具 - 快速生成PPT演示文稿
 *
 * 使用方法:
 *   node cli.js --title "我的演示" --slides "slides.json"
 *   node cli.js --from-dashboard
 */

import { PPTGenerator } from './index.js';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 加载环境变量
dotenv.config({ path: path.join(__dirname, '../../.env') });

// 解析命令行参数
function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    title: '演示文稿',
    slides: null,
    slidesFile: null,
    style: 'professional',
    fromDashboard: false,
    output: null,
    useNanoBanana: true,
    help: false
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    switch (arg) {
      case '-t':
      case '--title':
        options.title = args[++i];
        break;
      case '-s':
      case '--slides':
        options.slidesFile = args[++i];
        break;
      case '--style':
        options.style = args[++i];
        break;
      case '-d':
      case '--from-dashboard':
        options.fromDashboard = true;
        break;
      case '-o':
      case '--output':
        options.output = args[++i];
        break;
      case '--no-nanobanana':
        options.useNanoBanana = false;
        break;
      case '-h':
      case '--help':
        options.help = true;
        break;
    }
  }

  return options;
}

// 显示帮助信息
function showHelp() {
  console.log(`
╔═══════════════════════════════════════════════════════════════╗
║          🍌 PPT Generator - Nano Banana Pro Skill             ║
╚═══════════════════════════════════════════════════════════════╝

使用方法:
  node cli.js [options]

选项:
  -t, --title <标题>       设置PPT标题 (默认: "演示文稿")
  -s, --slides <文件>      从JSON文件加载幻灯片内容
  --style <风格>           幻灯片风格: professional, minimal, vibrant
  -d, --from-dashboard     从仪表盘数据自动生成PPT
  -o, --output <文件名>    输出文件名
  --no-nanobanana         禁用Nano Banana Pro图像生成
  -h, --help              显示帮助信息

环境变量:
  GEMINI_API_KEY          Gemini API密钥 (必需)
  PPT_OUTPUT_DIR          输出目录 (默认: ./output)
  PPT_DEFAULT_STYLE       默认风格 (默认: professional)

示例:
  # 从JSON文件生成PPT
  node cli.js -t "季度报告" -s slides.json --style professional

  # 从仪表盘数据生成
  node cli.js --from-dashboard -t "格力销售分析"

  # 使用模板方式(不使用Nano Banana Pro)
  node cli.js -t "简单演示" -s slides.json --no-nanobanana

幻灯片JSON格式:
  [
    {
      "title": "封面标题",
      "subtitle": "副标题"
    },
    {
      "title": "内容页",
      "bullets": ["要点1", "要点2", "要点3"]
    },
    "简单文本幻灯片"
  ]

更多信息: https://github.com/gree-dashboard/ppt-generator
`);
}

// 获取格力仪表盘数据
function getGreeDashboardData() {
  return {
    title: '格力抖音官方旗舰店 - 数据分析报告',
    kpis: [
      { label: '总销售额', value: '¥6.55亿' },
      { label: '总销量', value: '24.8万台' },
      { label: '整体客单价', value: '¥2,642' },
      { label: '业务周期', value: '40个月' }
    ],
    charts: [
      {
        title: '年度销售趋势分析',
        description: '2022-2025年销售额与销量变化趋势',
        highlights: [
          '2023年同比增长546.3%，实现爆发式增长',
          '2024年销售额达到峰值3.37亿元',
          '销量持续稳定增长，2025年预计8.3万台'
        ]
      },
      {
        title: '渠道分布分析',
        description: '直播、商品卡、其他渠道的销售占比',
        highlights: [
          '直播渠道贡献76%销售额(4.98亿)',
          '直播渠道销量占比达76%(18.87万台)',
          '商品卡渠道贡献16.5%销售额'
        ]
      },
      {
        title: 'Top 15 品类销售排名',
        description: '各空调型号销售业绩对比',
        highlights: [
          '格力王者空调柜机排名第一(8520万)',
          '云逸、云恬空调挂机位列二三名',
          '挂机产品占据销售榜主导地位'
        ]
      },
      {
        title: '月度季节性分析',
        description: '月度销售额的季节性波动规律',
        highlights: [
          '4-6月为销售旺季，6月达峰值4656万',
          '1-2月和8月为销售淡季',
          '明显的空调行业季节性特征'
        ]
      }
    ],
    insights: [
      '直播电商是核心增长引擎，贡献超75%业绩',
      '挂机空调是主力产品，覆盖多价位段',
      '季节性明显，需提前布局旺季备货',
      '2023年爆发式增长验证了直播电商模式的成功',
      '建议加强商品卡渠道建设，实现渠道多元化'
    ]
  };
}

// 主函数
async function main() {
  const options = parseArgs();

  if (options.help) {
    showHelp();
    process.exit(0);
  }

  // 检查API密钥
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.error('❌ 错误: 未设置 GEMINI_API_KEY 环境变量');
    console.error('请创建 .env 文件并添加: GEMINI_API_KEY=your_api_key');
    console.error('获取API密钥: https://aistudio.google.com/app/apikey');
    process.exit(1);
  }

  console.log('╔═══════════════════════════════════════════════════════════════╗');
  console.log('║          🍌 PPT Generator - Nano Banana Pro Skill             ║');
  console.log('╚═══════════════════════════════════════════════════════════════╝');

  try {
    const generator = new PPTGenerator(apiKey, {
      outputDir: process.env.PPT_OUTPUT_DIR || path.join(__dirname, '../../output'),
      defaultStyle: process.env.PPT_DEFAULT_STYLE || options.style
    });

    let outputPath;

    if (options.fromDashboard) {
      // 从仪表盘数据生成
      console.log('\n📊 从仪表盘数据生成PPT...');
      const dashboardData = getGreeDashboardData();
      dashboardData.title = options.title || dashboardData.title;
      outputPath = await generator.generateFromDashboard(dashboardData);
    } else if (options.slidesFile) {
      // 从JSON文件加载幻灯片
      console.log(`\n📄 从文件加载幻灯片: ${options.slidesFile}`);
      const slidesContent = await fs.readFile(options.slidesFile, 'utf-8');
      const slides = JSON.parse(slidesContent);
      outputPath = await generator.generate({
        title: options.title,
        slides,
        style: options.style,
        useNanoBanana: options.useNanoBanana,
        outputFileName: options.output
      });
    } else {
      // 使用示例幻灯片
      console.log('\n📝 使用示例幻灯片内容...');
      const slides = [
        {
          title: options.title,
          subtitle: '使用 Nano Banana Pro 生成'
        },
        {
          title: '主要内容',
          bullets: [
            '这是一个示例PPT',
            '由 Nano Banana Pro (Gemini 3 Pro) 生成',
            '支持中文文本渲染',
            '可自定义多种风格'
          ]
        },
        {
          title: '感谢观看',
          content: '由 Gree Dashboard PPT Skill 生成'
        }
      ];
      outputPath = await generator.generate({
        title: options.title,
        slides,
        style: options.style,
        useNanoBanana: options.useNanoBanana,
        outputFileName: options.output
      });
    }

    console.log('\n🎉 完成!');
    console.log(`📁 输出文件: ${outputPath}`);

  } catch (error) {
    console.error('\n❌ 生成失败:', error.message);
    if (process.env.DEBUG) {
      console.error(error.stack);
    }
    process.exit(1);
  }
}

main();
