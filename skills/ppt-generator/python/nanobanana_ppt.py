#!/usr/bin/env python3
"""
Nano Banana Pro PPT Generator
==============================
使用 Nano Banana Pro (Gemini 3 Pro) 生成高质量PPT演示文稿

深度参考 Manus AI 的实现方式:
1. 将每张幻灯片作为完整图像生成
2. 使用精心设计的prompt确保文本清晰可读
3. 支持多种风格和主题
4. 批量并行生成提升效率

Author: Gree Dashboard Team
"""

import os
import io
import json
import base64
import asyncio
from pathlib import Path
from typing import List, Dict, Any, Optional, Union
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum

from google import genai
from google.genai import types
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from dotenv import load_dotenv
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn

# 加载环境变量
load_dotenv()

console = Console()


class SlideStyle(Enum):
    """幻灯片风格枚举 - 参考Manus的风格系统"""
    PROFESSIONAL = "professional"      # 专业商务
    MINIMAL = "minimal"               # 极简白色
    VIBRANT = "vibrant"               # 活力渐变
    DARK_TECH = "dark_tech"           # 深色科技
    GRADIENT_BLUE = "gradient_blue"   # 蓝色渐变
    CORPORATE = "corporate"           # 企业标准


@dataclass
class SlideContent:
    """幻灯片内容数据结构"""
    title: str = ""
    subtitle: str = ""
    content: str = ""
    bullets: List[str] = field(default_factory=list)
    image_path: Optional[str] = None
    chart_data: Optional[Dict] = None
    layout: str = "title_content"  # title_only, title_content, two_column, chart
    notes: str = ""


@dataclass
class PPTConfig:
    """PPT配置"""
    title: str = "演示文稿"
    author: str = "Gree Dashboard"
    style: SlideStyle = SlideStyle.PROFESSIONAL
    width: float = 13.333  # 16:9 宽度(英寸)
    height: float = 7.5    # 16:9 高度(英寸)
    use_nanobanana: bool = True
    output_dir: str = "./output"
    language: str = "zh-CN"


# Manus风格的Prompt模板 - 这是他们PPT生成能力强大的核心
MANUS_STYLE_PROMPTS = {
    SlideStyle.PROFESSIONAL: """
Create a professional business presentation slide with these specifications:

**Visual Design:**
- Background: Deep navy blue gradient (#0F3460 to #1A1A2E)
- Primary text: Bright cyan (#00D4FF) for titles
- Secondary text: Clean white (#FFFFFF) for body
- Accent color: Electric green (#00FF88) for highlights
- Modern sans-serif typography, clean and readable

**Layout Rules:**
- Title positioned at top-left with bold weight
- Clear visual hierarchy with proper spacing
- Adequate padding (min 5% margins)
- Text must be crystal clear and anti-aliased

**Content:**
{content}

**Critical Requirements:**
- Text must be perfectly legible, no blur or artifacts
- Chinese characters must render correctly with proper fonts
- Maintain professional corporate aesthetic
- Resolution: 1920x1080 pixels
""",

    SlideStyle.MINIMAL: """
Create a minimalist presentation slide with clean aesthetics:

**Visual Design:**
- Background: Pure white (#FFFFFF)
- Primary text: Dark charcoal (#1A1A2E) for titles
- Secondary text: Medium gray (#666666) for body
- Accent: Single color highlight (#0F3460)
- Generous whitespace, breathing room

**Layout Rules:**
- Centered or left-aligned content
- Large, bold typography for impact
- Minimal decorative elements
- Focus on content clarity

**Content:**
{content}

**Critical Requirements:**
- Extreme clarity and readability
- Chinese text with elegant font rendering
- Clean, uncluttered composition
- Resolution: 1920x1080 pixels
""",

    SlideStyle.VIBRANT: """
Create a vibrant, energetic presentation slide:

**Visual Design:**
- Background: Dark gradient with colorful accents (#1A1A2E base)
- Primary text: Bright yellow (#FFD93D) for titles
- Secondary text: Coral red (#FF6B6B) for emphasis
- Body text: White (#FFFFFF)
- Dynamic, modern feel with subtle glow effects

**Layout Rules:**
- Bold, attention-grabbing typography
- Asymmetric layouts welcome
- Color blocks for emphasis
- High contrast for readability

**Content:**
{content}

**Critical Requirements:**
- Vivid colors but still professional
- Text remains highly legible
- Chinese characters properly rendered
- Resolution: 1920x1080 pixels
""",

    SlideStyle.DARK_TECH: """
Create a dark technology-themed presentation slide:

**Visual Design:**
- Background: Near-black with subtle tech patterns (#0a0a0f)
- Primary text: Neon cyan (#00f5ff) for titles
- Secondary text: Electric purple (#a855f7) for accents
- Body text: Light gray (#e0e0e0)
- Subtle grid or circuit patterns in background

**Layout Rules:**
- Futuristic, cutting-edge aesthetic
- Glowing text effects (subtle)
- Clean geometric shapes
- Tech-inspired iconography

**Content:**
{content}

**Critical Requirements:**
- High-tech appearance
- All text clearly readable
- Chinese support maintained
- Resolution: 1920x1080 pixels
""",

    SlideStyle.GRADIENT_BLUE: """
Create a elegant blue gradient presentation slide:

**Visual Design:**
- Background: Smooth blue gradient (#667eea to #764ba2)
- Primary text: White (#FFFFFF) for titles
- Secondary text: Light cyan (#e0f7ff)
- Accent: Golden yellow (#ffd700)
- Soft, professional gradients

**Layout Rules:**
- Elegant, flowing design
- Smooth transitions
- Professional corporate look
- Balanced composition

**Content:**
{content}

**Critical Requirements:**
- Premium, polished appearance
- Perfect text legibility on gradient
- Chinese text beautifully rendered
- Resolution: 1920x1080 pixels
""",

    SlideStyle.CORPORATE: """
Create a standard corporate presentation slide:

**Visual Design:**
- Background: White with subtle gray accents (#FAFAFA)
- Primary text: Corporate blue (#003366) for titles
- Secondary text: Dark gray (#333333)
- Accent: Brand orange (#FF6600) or teal (#008080)
- Clean, trustworthy aesthetic

**Layout Rules:**
- Standard business formatting
- Logo placement area (top-right)
- Footer with page numbers
- Conservative, reliable design

**Content:**
{content}

**Critical Requirements:**
- Professional and trustworthy
- Excellent readability
- Standard business proportions
- Resolution: 1920x1080 pixels
"""
}


class NanoBananaClient:
    """
    Nano Banana Pro API 客户端

    深度优化的实现，参考Manus的技术方案:
    1. 使用最新的 Gemini 3 Pro Image 模型
    2. 精确控制输出格式和质量
    3. 异步批量处理提升效率
    """

    # 模型优先级列表 - 按性能排序
    MODELS = [
        "gemini-2.0-flash-exp-image-generation",  # 当前最佳
        "gemini-2.0-flash-exp",                    # 备选
        "gemini-2.5-flash-preview-04-17",         # 稳定版
    ]

    def __init__(self, api_key: str):
        """初始化客户端"""
        self.api_key = api_key
        self.client = genai.Client(api_key=api_key)
        self.model = self._select_best_model()
        console.print(f"[green]✓[/green] Nano Banana Pro 客户端初始化完成")
        console.print(f"  使用模型: [cyan]{self.model}[/cyan]")

    def _select_best_model(self) -> str:
        """选择最佳可用模型"""
        # 目前使用第一个模型，后续可以添加模型探测逻辑
        return self.MODELS[0]

    def _build_slide_prompt(
        self,
        content: SlideContent,
        style: SlideStyle,
        slide_number: int,
        total_slides: int
    ) -> str:
        """
        构建幻灯片生成提示词

        Manus的核心技术之一：精确的prompt工程
        """
        # 组装内容描述
        content_parts = []

        if content.title:
            content_parts.append(f"**Title:** {content.title}")

        if content.subtitle:
            content_parts.append(f"**Subtitle:** {content.subtitle}")

        if content.content:
            content_parts.append(f"**Body Text:**\n{content.content}")

        if content.bullets:
            bullets_text = "\n".join(f"• {b}" for b in content.bullets)
            content_parts.append(f"**Key Points:**\n{bullets_text}")

        content_str = "\n\n".join(content_parts)

        # 添加幻灯片位置信息
        position_hint = ""
        if slide_number == 1:
            position_hint = "\n**Slide Type:** Title/Cover slide - make it impactful and memorable"
        elif slide_number == total_slides:
            position_hint = "\n**Slide Type:** Closing slide - include 'Thank You' or call-to-action feel"
        else:
            position_hint = f"\n**Slide Type:** Content slide ({slide_number} of {total_slides})"

        # 获取风格模板
        template = MANUS_STYLE_PROMPTS.get(style, MANUS_STYLE_PROMPTS[SlideStyle.PROFESSIONAL])

        # 组装最终prompt
        final_prompt = template.format(content=content_str + position_hint)

        return final_prompt

    async def generate_slide_image(
        self,
        content: SlideContent,
        style: SlideStyle = SlideStyle.PROFESSIONAL,
        slide_number: int = 1,
        total_slides: int = 1
    ) -> Optional[bytes]:
        """
        生成单张幻灯片图像

        Returns:
            图像字节数据，失败返回None
        """
        prompt = self._build_slide_prompt(content, style, slide_number, total_slides)

        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["Text", "Image"],
                    temperature=0.7,
                )
            )

            # 提取图像数据
            if response.candidates and response.candidates[0].content.parts:
                for part in response.candidates[0].content.parts:
                    if hasattr(part, 'inline_data') and part.inline_data:
                        image_data = part.inline_data.data
                        if isinstance(image_data, str):
                            return base64.b64decode(image_data)
                        return image_data

            console.print(f"[yellow]⚠[/yellow] 幻灯片 {slide_number} 未生成图像")
            return None

        except Exception as e:
            console.print(f"[red]✗[/red] 幻灯片 {slide_number} 生成失败: {e}")
            return None

    async def generate_slides_batch(
        self,
        slides: List[SlideContent],
        style: SlideStyle = SlideStyle.PROFESSIONAL,
        max_concurrent: int = 3
    ) -> List[Optional[bytes]]:
        """
        批量生成幻灯片图像

        使用异步并发提升效率，但限制并发数避免API限流
        """
        total = len(slides)
        results = [None] * total

        semaphore = asyncio.Semaphore(max_concurrent)

        async def generate_with_limit(idx: int, content: SlideContent):
            async with semaphore:
                # 添加延迟避免速率限制
                if idx > 0:
                    await asyncio.sleep(1.0)
                result = await self.generate_slide_image(
                    content, style, idx + 1, total
                )
                results[idx] = result

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            console=console
        ) as progress:
            task = progress.add_task(
                f"[cyan]生成 {total} 张幻灯片...",
                total=total
            )

            tasks = []
            for i, slide in enumerate(slides):
                tasks.append(generate_with_limit(i, slide))

            # 逐个完成以更新进度
            for coro in asyncio.as_completed(tasks):
                await coro
                progress.advance(task)

        return results


class PPTGenerator:
    """
    PPT 生成器主类

    支持两种模式:
    1. Nano Banana Pro 图像模式 - 生成高质量视觉幻灯片
    2. 模板模式 - 使用python-pptx生成可编辑幻灯片
    """

    def __init__(self, api_key: str, config: Optional[PPTConfig] = None):
        """初始化生成器"""
        self.api_key = api_key
        self.config = config or PPTConfig()
        self.nanobanana = NanoBananaClient(api_key) if self.config.use_nanobanana else None

        # 确保输出目录存在
        Path(self.config.output_dir).mkdir(parents=True, exist_ok=True)

    async def generate(
        self,
        slides: List[Union[SlideContent, Dict, str]],
        output_filename: Optional[str] = None
    ) -> str:
        """
        生成PPT文件

        Args:
            slides: 幻灯片内容列表
            output_filename: 输出文件名

        Returns:
            生成的文件路径
        """
        console.print(f"\n[bold cyan]{'='*60}[/bold cyan]")
        console.print(f"[bold]🍌 Nano Banana Pro PPT Generator[/bold]")
        console.print(f"[bold cyan]{'='*60}[/bold cyan]\n")

        console.print(f"📊 标题: [bold]{self.config.title}[/bold]")
        console.print(f"🎨 风格: [cyan]{self.config.style.value}[/cyan]")
        console.print(f"📄 幻灯片数量: [green]{len(slides)}[/green]")
        console.print(f"🍌 Nano Banana Pro: [{'green' if self.config.use_nanobanana else 'yellow'}]{'启用' if self.config.use_nanobanana else '禁用'}[/]\n")

        # 标准化幻灯片内容
        normalized_slides = self._normalize_slides(slides)

        # 创建演示文稿
        prs = Presentation()
        prs.slide_width = Inches(self.config.width)
        prs.slide_height = Inches(self.config.height)

        if self.config.use_nanobanana and self.nanobanana:
            # 使用 Nano Banana Pro 生成图像式幻灯片
            images = await self.nanobanana.generate_slides_batch(
                normalized_slides,
                self.config.style
            )

            for i, (slide_content, image_data) in enumerate(zip(normalized_slides, images)):
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

                if image_data:
                    # 添加生成的图像作为背景
                    image_stream = io.BytesIO(image_data)
                    slide.shapes.add_picture(
                        image_stream,
                        Inches(0), Inches(0),
                        width=prs.slide_width,
                        height=prs.slide_height
                    )
                    console.print(f"[green]✓[/green] 幻灯片 {i+1}: 图像已添加")
                else:
                    # 回退到模板模式
                    self._add_template_slide(slide, slide_content)
                    console.print(f"[yellow]⚠[/yellow] 幻灯片 {i+1}: 使用模板回退")
        else:
            # 使用模板模式
            for i, slide_content in enumerate(normalized_slides):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                self._add_template_slide(slide, slide_content)
                console.print(f"[green]✓[/green] 幻灯片 {i+1}: 模板生成完成")

        # 生成输出文件名
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_title = "".join(c if c.isalnum() or c in ('_', '-', ' ') else '_'
                                  for c in self.config.title)
            output_filename = f"{safe_title}_{timestamp}.pptx"

        output_path = Path(self.config.output_dir) / output_filename
        prs.save(str(output_path))

        console.print(f"\n[bold green]✅ PPT生成成功![/bold green]")
        console.print(f"📁 文件路径: [cyan]{output_path}[/cyan]\n")

        return str(output_path)

    def _normalize_slides(self, slides: List[Union[SlideContent, Dict, str]]) -> List[SlideContent]:
        """标准化幻灯片内容格式"""
        normalized = []

        for slide in slides:
            if isinstance(slide, SlideContent):
                normalized.append(slide)
            elif isinstance(slide, dict):
                normalized.append(SlideContent(
                    title=slide.get('title', ''),
                    subtitle=slide.get('subtitle', ''),
                    content=slide.get('content', ''),
                    bullets=slide.get('bullets', []),
                    layout=slide.get('layout', 'title_content'),
                    notes=slide.get('notes', '')
                ))
            elif isinstance(slide, str):
                normalized.append(SlideContent(title=slide))
            else:
                console.print(f"[yellow]⚠[/yellow] 跳过无效的幻灯片内容: {type(slide)}")

        return normalized

    def _add_template_slide(self, slide, content: SlideContent):
        """添加模板式幻灯片（回退方案）"""
        # 根据风格获取颜色配置
        colors = self._get_style_colors()

        # 添加背景
        background = slide.shapes.add_shape(
            1,  # 矩形
            Inches(0), Inches(0),
            Inches(self.config.width), Inches(self.config.height)
        )
        background.fill.solid()
        background.fill.fore_color.rgb = RgbColor.from_string(colors['background'])
        background.line.fill.background()

        y_offset = 0.5

        # 标题
        if content.title:
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(y_offset),
                Inches(self.config.width - 1), Inches(0.8)
            )
            title_frame = title_box.text_frame
            title_para = title_frame.paragraphs[0]
            title_para.text = content.title
            title_para.font.size = Pt(36)
            title_para.font.bold = True
            title_para.font.color.rgb = RgbColor.from_string(colors['title'])
            y_offset += 0.9

        # 副标题
        if content.subtitle:
            sub_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(y_offset),
                Inches(self.config.width - 1), Inches(0.5)
            )
            sub_frame = sub_box.text_frame
            sub_para = sub_frame.paragraphs[0]
            sub_para.text = content.subtitle
            sub_para.font.size = Pt(18)
            sub_para.font.color.rgb = RgbColor.from_string(colors['subtitle'])
            y_offset += 0.6

        # 正文或要点
        if content.content or content.bullets:
            body_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(y_offset),
                Inches(self.config.width - 1), Inches(self.config.height - y_offset - 0.5)
            )
            body_frame = body_box.text_frame
            body_frame.word_wrap = True

            if content.content:
                p = body_frame.paragraphs[0]
                p.text = content.content
                p.font.size = Pt(14)
                p.font.color.rgb = RgbColor.from_string(colors['body'])

            if content.bullets:
                for i, bullet in enumerate(content.bullets):
                    if i == 0 and not content.content:
                        p = body_frame.paragraphs[0]
                    else:
                        p = body_frame.add_paragraph()
                    p.text = f"• {bullet}"
                    p.font.size = Pt(14)
                    p.font.color.rgb = RgbColor.from_string(colors['body'])
                    p.level = 0

    def _get_style_colors(self) -> Dict[str, str]:
        """获取风格对应的颜色配置"""
        style_colors = {
            SlideStyle.PROFESSIONAL: {
                'background': '0F3460',
                'title': '00D4FF',
                'subtitle': 'FFFFFF',
                'body': 'E0E0E0',
                'accent': '00FF88'
            },
            SlideStyle.MINIMAL: {
                'background': 'FFFFFF',
                'title': '1A1A2E',
                'subtitle': '666666',
                'body': '333333',
                'accent': '0F3460'
            },
            SlideStyle.VIBRANT: {
                'background': '1A1A2E',
                'title': 'FFD93D',
                'subtitle': 'FF6B6B',
                'body': 'FFFFFF',
                'accent': '00D4FF'
            },
            SlideStyle.DARK_TECH: {
                'background': '0A0A0F',
                'title': '00F5FF',
                'subtitle': 'A855F7',
                'body': 'E0E0E0',
                'accent': '00FF88'
            },
            SlideStyle.GRADIENT_BLUE: {
                'background': '667EEA',
                'title': 'FFFFFF',
                'subtitle': 'E0F7FF',
                'body': 'FFFFFF',
                'accent': 'FFD700'
            },
            SlideStyle.CORPORATE: {
                'background': 'FAFAFA',
                'title': '003366',
                'subtitle': '666666',
                'body': '333333',
                'accent': 'FF6600'
            }
        }
        return style_colors.get(self.config.style, style_colors[SlideStyle.PROFESSIONAL])

    @classmethod
    async def from_dashboard_data(
        cls,
        api_key: str,
        dashboard_data: Dict[str, Any],
        style: SlideStyle = SlideStyle.PROFESSIONAL
    ) -> str:
        """
        从仪表盘数据生成PPT

        这是Manus的核心功能之一：自动将数据转换为演示文稿
        """
        title = dashboard_data.get('title', '数据分析报告')
        kpis = dashboard_data.get('kpis', [])
        charts = dashboard_data.get('charts', [])
        insights = dashboard_data.get('insights', [])

        slides = []

        # 封面
        slides.append(SlideContent(
            title=title,
            subtitle=f"报告生成时间: {datetime.now().strftime('%Y年%m月%d日')}",
            layout='title_only'
        ))

        # KPI概览
        if kpis:
            kpi_bullets = [f"{kpi['label']}: {kpi['value']}" for kpi in kpis]
            slides.append(SlideContent(
                title="关键业绩指标 (KPIs)",
                bullets=kpi_bullets
            ))

        # 图表分析页
        for chart in charts:
            slides.append(SlideContent(
                title=chart.get('title', ''),
                content=chart.get('description', ''),
                bullets=chart.get('highlights', [])
            ))

        # 洞察总结
        if insights:
            slides.append(SlideContent(
                title="核心洞察与建议",
                bullets=insights
            ))

        # 结尾
        slides.append(SlideContent(
            title="感谢观看",
            subtitle="Thank You",
            content="由 Gree Dashboard PPT Skill 生成\n使用 Nano Banana Pro (Gemini 3 Pro) 技术"
        ))

        config = PPTConfig(
            title=title,
            style=style,
            use_nanobanana=True
        )

        generator = cls(api_key, config)
        return await generator.generate(slides)


def load_slides_from_json(json_path: str) -> List[Dict]:
    """从JSON文件加载幻灯片内容"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


async def main():
    """主入口函数 - 示例用法"""
    import typer
    from typing_extensions import Annotated

    app = typer.Typer(
        name="nanobanana-ppt",
        help="🍌 Nano Banana Pro PPT Generator - 使用 Gemini 3 Pro 生成高质量演示文稿"
    )

    @app.command()
    def generate(
        title: Annotated[str, typer.Option("--title", "-t", help="PPT标题")] = "演示文稿",
        slides_file: Annotated[Optional[str], typer.Option("--slides", "-s", help="幻灯片JSON文件路径")] = None,
        style: Annotated[str, typer.Option("--style", help="风格: professional, minimal, vibrant, dark_tech")] = "professional",
        from_dashboard: Annotated[bool, typer.Option("--from-dashboard", "-d", help="从仪表盘数据生成")] = False,
        output: Annotated[Optional[str], typer.Option("--output", "-o", help="输出文件名")] = None,
        no_nanobanana: Annotated[bool, typer.Option("--no-nanobanana", help="禁用Nano Banana Pro图像生成")] = False,
    ):
        """生成PPT演示文稿"""
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            console.print("[red]错误: 未设置 GEMINI_API_KEY 环境变量[/red]")
            console.print("请在 .env 文件中添加: GEMINI_API_KEY=your_api_key")
            raise typer.Exit(1)

        try:
            style_enum = SlideStyle(style)
        except ValueError:
            console.print(f"[red]错误: 无效的风格 '{style}'[/red]")
            console.print(f"可用风格: {', '.join(s.value for s in SlideStyle)}")
            raise typer.Exit(1)

        async def run():
            if from_dashboard:
                # 使用格力仪表盘数据
                dashboard_data = {
                    'title': title if title != "演示文稿" else '格力抖音官方旗舰店 - 数据分析报告',
                    'kpis': [
                        {'label': '总销售额', 'value': '¥6.55亿'},
                        {'label': '总销量', 'value': '24.8万台'},
                        {'label': '整体客单价', 'value': '¥2,642'},
                        {'label': '业务周期', 'value': '40个月'}
                    ],
                    'charts': [
                        {
                            'title': '年度销售趋势分析',
                            'description': '2022-2025年销售额与销量变化趋势',
                            'highlights': [
                                '2023年同比增长546.3%，实现爆发式增长',
                                '2024年销售额达到峰值3.37亿元',
                                '销量持续稳定增长，2025年预计8.3万台'
                            ]
                        },
                        {
                            'title': '渠道分布分析',
                            'description': '直播、商品卡、其他渠道的销售占比',
                            'highlights': [
                                '直播渠道贡献76%销售额(4.98亿)',
                                '直播渠道销量占比达76%(18.87万台)',
                                '商品卡渠道贡献16.5%销售额'
                            ]
                        },
                        {
                            'title': '月度季节性分析',
                            'description': '月度销售额的季节性波动规律',
                            'highlights': [
                                '4-6月为销售旺季，6月达峰值4656万',
                                '1-2月和8月为销售淡季',
                                '明显的空调行业季节性特征'
                            ]
                        }
                    ],
                    'insights': [
                        '直播电商是核心增长引擎，贡献超75%业绩',
                        '挂机空调是主力产品，覆盖多价位段',
                        '季节性明显，需提前布局旺季备货',
                        '建议加强商品卡渠道建设，实现渠道多元化'
                    ]
                }
                return await PPTGenerator.from_dashboard_data(api_key, dashboard_data, style_enum)

            elif slides_file:
                slides = load_slides_from_json(slides_file)
                config = PPTConfig(
                    title=title,
                    style=style_enum,
                    use_nanobanana=not no_nanobanana
                )
                generator = PPTGenerator(api_key, config)
                return await generator.generate(slides, output)

            else:
                # 使用示例幻灯片
                slides = [
                    SlideContent(title=title, subtitle="使用 Nano Banana Pro 生成"),
                    SlideContent(
                        title="主要内容",
                        bullets=[
                            "这是一个示例PPT",
                            "由 Nano Banana Pro (Gemini 3 Pro) 生成",
                            "支持中文文本渲染",
                            "可自定义多种风格"
                        ]
                    ),
                    SlideContent(title="感谢观看", content="由 Gree Dashboard PPT Skill 生成")
                ]
                config = PPTConfig(
                    title=title,
                    style=style_enum,
                    use_nanobanana=not no_nanobanana
                )
                generator = PPTGenerator(api_key, config)
                return await generator.generate(slides, output)

        output_path = asyncio.run(run())
        console.print(f"\n[bold green]🎉 完成![/bold green] 文件: {output_path}")

    app()


if __name__ == "__main__":
    main()
