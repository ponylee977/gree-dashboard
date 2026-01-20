#!/usr/bin/env python3
"""
Manus Clone PPT Generator - 深度克隆版
=========================================

双模型协作架构:
- Gemini 3 Pro: 负责PPT大纲、框架、写作逻辑
- Claude Opus 4.5: 负责细节内容填充、润色、质检

核心特性:
- 8K分辨率 (7680x4320) 确保文字清晰
- 硬编码内容控制，防止幻觉
- 深度克隆Manus的参数和功能
- 多层质检机制

Author: Gree Dashboard Team
"""

import os
import io
import json
import base64
import asyncio
import httpx
from pathlib import Path
from typing import List, Dict, Any, Optional, Union, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
import re

from google import genai
from google.genai import types
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from dotenv import load_dotenv
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn
from rich.panel import Panel
from rich.table import Table

load_dotenv()
console = Console()


# ============================================================================
# 硬编码配置 - Manus风格参数克隆
# ============================================================================

class ManusConfig:
    """Manus深度克隆配置 - 硬编码防止幻觉"""

    # 8K分辨率配置 - 确保文字清晰
    RESOLUTION_8K = {
        "width": 7680,
        "height": 4320,
        "aspect_ratio": "16:9",
        "dpi": 300
    }

    # 4K备选配置
    RESOLUTION_4K = {
        "width": 3840,
        "height": 2160,
        "aspect_ratio": "16:9",
        "dpi": 150
    }

    # PPT尺寸 (英寸)
    SLIDE_SIZE = {
        "width": 13.333,
        "height": 7.5
    }

    # 字体配置 - 确保中文渲染
    FONTS = {
        "chinese_primary": "Microsoft YaHei",
        "chinese_fallback": "PingFang SC",
        "english": "Arial",
        "monospace": "Consolas"
    }

    # 最小字号限制 (pt) - 8K下保证可读性
    MIN_FONT_SIZES = {
        "title": 72,      # 8K下标题最小72pt
        "subtitle": 48,   # 副标题48pt
        "body": 36,       # 正文36pt
        "caption": 28     # 说明文字28pt
    }

    # API速率限制
    RATE_LIMITS = {
        "gemini_rpm": 15,        # Gemini每分钟请求数
        "claude_rpm": 50,        # Claude每分钟请求数
        "delay_between_slides": 2.0,  # 幻灯片间延迟(秒)
        "retry_attempts": 3,     # 重试次数
        "retry_delay": 5.0       # 重试延迟(秒)
    }

    # 内容限制 - 防止幻觉
    CONTENT_LIMITS = {
        "max_title_length": 30,       # 标题最大字符数
        "max_subtitle_length": 50,    # 副标题最大字符数
        "max_bullets_per_slide": 6,   # 每页最多要点数
        "max_bullet_length": 80,      # 每个要点最大字符数
        "max_body_length": 300,       # 正文最大字符数
        "max_slides": 20              # 最大幻灯片数
    }

    # 质检规则
    QUALITY_CHECKS = {
        "require_title": True,           # 必须有标题
        "no_empty_slides": True,         # 不允许空幻灯片
        "check_chinese_encoding": True,  # 检查中文编码
        "validate_numbers": True,        # 验证数字格式
        "check_contrast": True           # 检查对比度
    }


# ============================================================================
# Manus风格模板 - 8K优化版
# ============================================================================

MANUS_8K_PROMPTS = {
    "professional": """
Generate an 8K resolution (7680x4320 pixels) professional business presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Output resolution: EXACTLY 7680x4320 pixels
- All text must be rendered at HIGH DPI (300 DPI equivalent)
- Chinese characters: Use Microsoft YaHei or similar, minimum 72pt for titles
- Numbers and English: Use Arial or similar sans-serif, crisp anti-aliasing
- NO blur, NO artifacts, NO compression artifacts on text

**Visual Design:**
- Background: Deep navy blue gradient (#0F3460 to #1A1A2E)
- Title: Bright cyan (#00D4FF), 72-96pt, bold
- Subtitle: White (#FFFFFF), 48-60pt
- Body text: Light gray (#E0E0E0), 36-48pt
- Accent elements: Electric green (#00FF88)

**Layout (8K optimized):**
- Title area: Top 15% of slide
- Content area: Middle 70%
- Footer area: Bottom 15%
- Side margins: 5% each side
- All text left-aligned for readability

**Content to render:**
{content}

**Slide context:**
{context}

**FINAL CHECK:**
- Text must be PIXEL PERFECT and CRYSTAL CLEAR
- Every Chinese character must be fully legible
- Every number must be sharp and distinct
- No visual artifacts or blur anywhere
""",

    "minimal_8k": """
Generate an 8K resolution (7680x4320 pixels) minimalist presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Output resolution: EXACTLY 7680x4320 pixels
- Ultra-high clarity text rendering
- Chinese: Microsoft YaHei/PingFang SC, minimum 72pt titles
- Maximum contrast between text and background

**Visual Design:**
- Background: Pure white (#FFFFFF)
- Title: Dark charcoal (#1A1A2E), 84pt, bold
- Body: Medium gray (#333333), 42pt
- Accent: Navy blue (#0F3460) for highlights
- Generous whitespace (40% of slide)

**Content to render:**
{content}

**Slide context:**
{context}

**Text must be ABSOLUTELY SHARP - no exceptions**
""",

    "dark_tech_8k": """
Generate an 8K resolution (7680x4320 pixels) dark technology presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Resolution: 7680x4320 pixels EXACTLY
- Neon text effects must remain READABLE
- Chinese characters: Extra attention to clarity on dark background

**Visual Design:**
- Background: Near-black (#0A0A0F) with subtle circuit patterns
- Title: Neon cyan (#00F5FF), 80pt, slight glow effect
- Subtitle: Electric purple (#A855F7), 52pt
- Body: Light gray (#E0E0E0), 40pt
- Grid/tech patterns at 10% opacity only

**Content to render:**
{content}

**Slide context:**
{context}

**WARNING: Dark backgrounds require EXTRA text clarity**
""",

    "gradient_premium_8k": """
Generate an 8K resolution (7680x4320 pixels) premium gradient presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Resolution: 7680x4320 pixels
- Text on gradients must have MAXIMUM contrast
- Consider adding subtle text shadow for legibility

**Visual Design:**
- Background: Smooth gradient (#667EEA → #764BA2)
- Title: Pure white (#FFFFFF) with subtle shadow, 78pt
- Subtitle: Light cyan (#E0F7FF), 50pt
- Body: White (#FFFFFF), 38pt
- Accent: Golden (#FFD700) for highlights

**Content to render:**
{content}

**Slide context:**
{context}

**Gradient backgrounds: ensure text POP with proper contrast**
""",

    "corporate_8k": """
Generate an 8K resolution (7680x4320 pixels) corporate presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Resolution: 7680x4320 pixels
- Conservative, trustworthy appearance
- Perfect for formal business settings

**Visual Design:**
- Background: Off-white (#FAFAFA)
- Title: Corporate blue (#003366), 76pt, bold
- Subtitle: Dark gray (#666666), 48pt
- Body: Charcoal (#333333), 36pt
- Accent: Professional orange (#FF6600)

**Layout:**
- Logo area: Top-right corner
- Page number: Bottom-right
- Company name: Bottom-left

**Content to render:**
{content}

**Slide context:**
{context}

**Corporate style: clean, professional, trustworthy**
"""
}


# ============================================================================
# Claude API客户端 - 内容填充和质检
# ============================================================================

class ClaudeClient:
    """Claude Opus 4.5 客户端 - 负责内容填充和质检"""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = "https://api.anthropic.com/v1"
        self.model = "claude-opus-4-5-20251101"
        console.print(f"[green]✓[/green] Claude Opus 4.5 客户端初始化完成")

    async def _request(self, messages: List[Dict], system: str = "", max_tokens: int = 4096) -> str:
        """发送请求到Claude API"""
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(
                f"{self.base_url}/messages",
                headers={
                    "x-api-key": self.api_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                },
                json={
                    "model": self.model,
                    "max_tokens": max_tokens,
                    "system": system,
                    "messages": messages
                }
            )
            response.raise_for_status()
            data = response.json()
            return data["content"][0]["text"]

    async def fill_slide_content(
        self,
        outline: Dict[str, Any],
        context: str,
        constraints: Dict[str, int]
    ) -> Dict[str, Any]:
        """
        填充幻灯片内容细节

        Args:
            outline: Gemini生成的大纲
            context: 上下文信息
            constraints: 内容限制

        Returns:
            填充后的完整内容
        """
        system_prompt = f"""你是一位专业的PPT内容撰写专家。你的任务是将大纲扩展为完整、专业的幻灯片内容。

**硬性约束 (必须严格遵守):**
- 标题最多 {constraints.get('max_title_length', 30)} 个字符
- 副标题最多 {constraints.get('max_subtitle_length', 50)} 个字符
- 每个要点最多 {constraints.get('max_bullet_length', 80)} 个字符
- 每页最多 {constraints.get('max_bullets_per_slide', 6)} 个要点
- 正文最多 {constraints.get('max_body_length', 300)} 个字符

**内容质量要求:**
1. 数据必须准确，不要编造数字
2. 使用专业、简洁的商务语言
3. 中文表达要地道流畅
4. 避免空洞的套话

**输出格式:** 严格的JSON格式"""

        user_message = f"""请将以下大纲扩展为完整的幻灯片内容:

**大纲:**
```json
{json.dumps(outline, ensure_ascii=False, indent=2)}
```

**背景上下文:**
{context}

请输出完整的JSON，包含以下字段:
- title: 标题
- subtitle: 副标题 (可选)
- bullets: 要点列表 (可选)
- content: 正文 (可选)
- notes: 演讲备注

只输出JSON，不要其他内容。"""

        try:
            response = await self._request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=2048
            )
            # 提取JSON
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                return json.loads(json_match.group())
            return outline
        except Exception as e:
            console.print(f"[yellow]⚠[/yellow] Claude填充失败: {e}")
            return outline

    async def quality_check(self, slide_content: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """
        质检幻灯片内容

        Returns:
            (是否通过, 问题列表)
        """
        issues = []
        config = ManusConfig.CONTENT_LIMITS

        # 标题检查
        title = slide_content.get('title', '')
        if not title and ManusConfig.QUALITY_CHECKS['require_title']:
            issues.append("缺少标题")
        elif len(title) > config['max_title_length']:
            issues.append(f"标题过长: {len(title)}/{config['max_title_length']}")

        # 副标题检查
        subtitle = slide_content.get('subtitle', '')
        if subtitle and len(subtitle) > config['max_subtitle_length']:
            issues.append(f"副标题过长: {len(subtitle)}/{config['max_subtitle_length']}")

        # 要点检查
        bullets = slide_content.get('bullets', [])
        if len(bullets) > config['max_bullets_per_slide']:
            issues.append(f"要点过多: {len(bullets)}/{config['max_bullets_per_slide']}")
        for i, bullet in enumerate(bullets):
            if len(bullet) > config['max_bullet_length']:
                issues.append(f"要点{i+1}过长: {len(bullet)}/{config['max_bullet_length']}")

        # 正文检查
        content = slide_content.get('content', '')
        if content and len(content) > config['max_body_length']:
            issues.append(f"正文过长: {len(content)}/{config['max_body_length']}")

        # 空内容检查
        if ManusConfig.QUALITY_CHECKS['no_empty_slides']:
            if not title and not bullets and not content:
                issues.append("幻灯片内容为空")

        # 中文编码检查
        if ManusConfig.QUALITY_CHECKS['check_chinese_encoding']:
            all_text = title + subtitle + content + ''.join(bullets)
            try:
                all_text.encode('utf-8')
            except UnicodeEncodeError:
                issues.append("包含无法编码的字符")

        return len(issues) == 0, issues

    async def fix_content_issues(self, slide_content: Dict[str, Any], issues: List[str]) -> Dict[str, Any]:
        """修复内容问题"""
        system_prompt = """你是PPT内容修复专家。请修复以下问题，确保内容符合约束条件。
只修复有问题的部分，保持其他内容不变。输出纯JSON格式。"""

        user_message = f"""原始内容:
```json
{json.dumps(slide_content, ensure_ascii=False, indent=2)}
```

发现的问题:
{chr(10).join(f'- {issue}' for issue in issues)}

约束条件:
- 标题最多30字符
- 副标题最多50字符
- 每个要点最多80字符
- 最多6个要点
- 正文最多300字符

请修复并输出JSON:"""

        try:
            response = await self._request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=2048
            )
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                return json.loads(json_match.group())
            return slide_content
        except Exception as e:
            console.print(f"[yellow]⚠[/yellow] 内容修复失败: {e}")
            return slide_content


# ============================================================================
# Gemini客户端 - 框架逻辑和图像生成
# ============================================================================

class GeminiClient:
    """Gemini 3 Pro 客户端 - 负责框架和图像生成"""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.client = genai.Client(api_key=api_key)
        self.text_model = "gemini-2.5-flash-preview-04-17"
        self.image_model = "gemini-2.0-flash-exp-image-generation"
        console.print(f"[green]✓[/green] Gemini 3 Pro 客户端初始化完成")

    async def generate_ppt_outline(
        self,
        topic: str,
        context: str,
        num_slides: int = 8
    ) -> List[Dict[str, Any]]:
        """
        生成PPT大纲框架

        Gemini负责:
        - 整体结构设计
        - 逻辑流程安排
        - 内容框架规划
        """
        prompt = f"""作为PPT架构师，请为以下主题设计演示文稿大纲:

**主题:** {topic}

**背景信息:**
{context}

**要求:**
- 生成 {num_slides} 张幻灯片的大纲
- 第1张: 封面
- 第2-{num_slides-1}张: 内容页
- 第{num_slides}张: 总结/感谢页

**每张幻灯片包含:**
- slide_number: 序号
- type: 类型 (cover/content/summary)
- title_hint: 标题提示 (10字以内)
- content_hints: 内容要点提示 (列表)
- visual_suggestion: 视觉建议

只输出JSON数组，不要其他内容:"""

        try:
            response = self.client.models.generate_content(
                model=self.text_model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.7,
                    max_output_tokens=4096
                )
            )
            text = response.text
            json_match = re.search(r'\[[\s\S]*\]', text)
            if json_match:
                return json.loads(json_match.group())
            return []
        except Exception as e:
            console.print(f"[red]✗[/red] Gemini大纲生成失败: {e}")
            return []

    async def generate_8k_slide_image(
        self,
        content: Dict[str, Any],
        style: str = "professional",
        slide_number: int = 1,
        total_slides: int = 1
    ) -> Optional[bytes]:
        """
        生成8K分辨率幻灯片图像

        关键: 使用8K优化的prompt确保文字清晰
        """
        # 获取8K优化的prompt模板
        template = MANUS_8K_PROMPTS.get(f"{style}_8k", MANUS_8K_PROMPTS.get(style, MANUS_8K_PROMPTS["professional"]))

        # 构建内容描述
        content_parts = []
        if content.get('title'):
            content_parts.append(f"**标题:** {content['title']}")
        if content.get('subtitle'):
            content_parts.append(f"**副标题:** {content['subtitle']}")
        if content.get('bullets'):
            bullets_text = '\n'.join(f"• {b}" for b in content['bullets'])
            content_parts.append(f"**要点:**\n{bullets_text}")
        if content.get('content'):
            content_parts.append(f"**正文:**\n{content['content']}")

        content_str = '\n\n'.join(content_parts)

        # 幻灯片上下文
        context_parts = []
        if slide_number == 1:
            context_parts.append("这是封面页，需要视觉冲击力")
        elif slide_number == total_slides:
            context_parts.append("这是结尾页，包含感谢或总结")
        else:
            context_parts.append(f"这是第 {slide_number}/{total_slides} 页内容页")

        context_str = '\n'.join(context_parts)

        # 组装最终prompt
        final_prompt = template.format(content=content_str, context=context_str)

        # 添加8K强制声明
        final_prompt += """

**ABSOLUTE REQUIREMENTS (DO NOT IGNORE):**
1. Resolution MUST be 7680x4320 pixels (8K)
2. All Chinese characters MUST be perfectly clear
3. All numbers MUST be sharp and readable
4. NO blur, NO artifacts, NO compression issues
5. Text quality is the TOP PRIORITY
"""

        try:
            response = self.client.models.generate_content(
                model=self.image_model,
                contents=final_prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["Text", "Image"],
                    temperature=0.8
                )
            )

            # 提取图像
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
            console.print(f"[red]✗[/red] 图像生成失败: {e}")
            return None


# ============================================================================
# Manus Clone PPT生成器主类
# ============================================================================

@dataclass
class ManusCloneConfig:
    """Manus Clone 配置"""
    title: str = "演示文稿"
    author: str = "Gree Dashboard"
    style: str = "professional"
    num_slides: int = 8
    use_8k: bool = True
    output_dir: str = "./output"
    enable_quality_check: bool = True
    enable_content_fill: bool = True


class ManusClonePPTGenerator:
    """
    Manus Clone PPT生成器

    双模型协作架构:
    1. Gemini: 框架设计 + 图像生成
    2. Claude: 内容填充 + 质量检查
    """

    def __init__(
        self,
        gemini_api_key: str,
        claude_api_key: str,
        config: Optional[ManusCloneConfig] = None
    ):
        self.config = config or ManusCloneConfig()
        self.gemini = GeminiClient(gemini_api_key)
        self.claude = ClaudeClient(claude_api_key)
        self.output_dir = Path(self.config.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        console.print(Panel.fit(
            "[bold cyan]Manus Clone PPT Generator[/bold cyan]\n"
            "双模型协作: Gemini (框架) + Claude (内容)\n"
            f"分辨率: {'8K (7680x4320)' if self.config.use_8k else '4K (3840x2160)'}",
            title="🍌 初始化完成"
        ))

    async def generate(
        self,
        topic: str,
        context: str = "",
        slides_data: Optional[List[Dict]] = None
    ) -> str:
        """
        生成PPT

        流程:
        1. Gemini生成大纲框架
        2. Claude填充内容细节
        3. Claude质检内容
        4. Gemini生成8K图像
        5. 组装PPT文件
        """
        console.print(f"\n[bold]🎯 主题:[/bold] {topic}")
        console.print(f"[bold]🎨 风格:[/bold] {self.config.style}")
        console.print(f"[bold]📊 页数:[/bold] {self.config.num_slides}")
        console.print()

        # Step 1: 生成或使用提供的幻灯片数据
        if slides_data:
            console.print("[cyan]使用提供的幻灯片数据...[/cyan]")
            outlines = slides_data
        else:
            console.print("[cyan]Step 1: Gemini 生成PPT框架...[/cyan]")
            outlines = await self.gemini.generate_ppt_outline(
                topic, context, self.config.num_slides
            )

            if not outlines:
                console.print("[red]框架生成失败，使用默认结构[/red]")
                outlines = self._get_default_outline(topic)

        console.print(f"[green]✓[/green] 获得 {len(outlines)} 张幻灯片框架\n")

        # Step 2: Claude填充内容
        filled_slides = []
        if self.config.enable_content_fill:
            console.print("[cyan]Step 2: Claude 填充内容细节...[/cyan]")

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                console=console
            ) as progress:
                task = progress.add_task("填充内容...", total=len(outlines))

                for i, outline in enumerate(outlines):
                    filled = await self.claude.fill_slide_content(
                        outline, context, ManusConfig.CONTENT_LIMITS
                    )
                    filled_slides.append(filled)
                    progress.advance(task)
                    await asyncio.sleep(0.5)  # 避免API限流
        else:
            filled_slides = outlines

        console.print(f"[green]✓[/green] 内容填充完成\n")

        # Step 3: 质量检查
        if self.config.enable_quality_check:
            console.print("[cyan]Step 3: Claude 质量检查...[/cyan]")
            checked_slides = []

            for i, slide in enumerate(filled_slides):
                passed, issues = await self.claude.quality_check(slide)

                if not passed:
                    console.print(f"  [yellow]⚠[/yellow] 幻灯片 {i+1} 发现问题: {', '.join(issues)}")
                    fixed = await self.claude.fix_content_issues(slide, issues)
                    checked_slides.append(fixed)
                else:
                    checked_slides.append(slide)

            filled_slides = checked_slides
            console.print(f"[green]✓[/green] 质检完成\n")

        # Step 4: 生成8K图像
        console.print("[cyan]Step 4: Gemini 生成8K幻灯片图像...[/cyan]")
        images = []
        total = len(filled_slides)

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            console=console
        ) as progress:
            task = progress.add_task("生成图像...", total=total)

            for i, slide in enumerate(filled_slides):
                image_data = await self.gemini.generate_8k_slide_image(
                    slide,
                    style=self.config.style,
                    slide_number=i + 1,
                    total_slides=total
                )
                images.append(image_data)
                progress.advance(task)

                # 延迟避免限流
                if i < total - 1:
                    await asyncio.sleep(ManusConfig.RATE_LIMITS['delay_between_slides'])

        success_count = sum(1 for img in images if img is not None)
        console.print(f"[green]✓[/green] 成功生成 {success_count}/{total} 张图像\n")

        # Step 5: 组装PPT
        console.print("[cyan]Step 5: 组装PPT文件...[/cyan]")
        output_path = await self._assemble_ppt(filled_slides, images)

        console.print(Panel.fit(
            f"[bold green]✅ PPT生成成功![/bold green]\n\n"
            f"📁 文件: {output_path}\n"
            f"📊 页数: {len(filled_slides)}\n"
            f"🎨 风格: {self.config.style}\n"
            f"📐 分辨率: {'8K' if self.config.use_8k else '4K'}",
            title="🍌 完成"
        ))

        return str(output_path)

    async def _assemble_ppt(
        self,
        slides: List[Dict],
        images: List[Optional[bytes]]
    ) -> Path:
        """组装PPT文件"""
        prs = Presentation()

        # 设置幻灯片尺寸
        prs.slide_width = Inches(ManusConfig.SLIDE_SIZE['width'])
        prs.slide_height = Inches(ManusConfig.SLIDE_SIZE['height'])

        for i, (slide_content, image_data) in enumerate(zip(slides, images)):
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

            if image_data:
                # 添加生成的图像
                image_stream = io.BytesIO(image_data)
                slide.shapes.add_picture(
                    image_stream,
                    Inches(0), Inches(0),
                    width=prs.slide_width,
                    height=prs.slide_height
                )
            else:
                # 回退到模板
                self._add_fallback_slide(slide, slide_content)

        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[^\w\u4e00-\u9fa5\-]', '_', self.config.title)
        filename = f"{safe_title}_{timestamp}.pptx"
        output_path = self.output_dir / filename

        prs.save(str(output_path))
        return output_path

    def _add_fallback_slide(self, slide, content: Dict):
        """添加回退模板幻灯片"""
        # 背景
        background = slide.shapes.add_shape(
            1, Inches(0), Inches(0),
            Inches(ManusConfig.SLIDE_SIZE['width']),
            Inches(ManusConfig.SLIDE_SIZE['height'])
        )
        background.fill.solid()
        background.fill.fore_color.rgb = RgbColor.from_string('0F3460')
        background.line.fill.background()

        y = 0.5

        # 标题
        if content.get('title'):
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(y),
                Inches(12), Inches(1)
            )
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.text = content['title']
            p.font.size = Pt(ManusConfig.MIN_FONT_SIZES['title'])
            p.font.bold = True
            p.font.color.rgb = RgbColor.from_string('00D4FF')
            y += 1.2

        # 副标题
        if content.get('subtitle'):
            sub_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(y),
                Inches(12), Inches(0.7)
            )
            tf = sub_box.text_frame
            p = tf.paragraphs[0]
            p.text = content['subtitle']
            p.font.size = Pt(ManusConfig.MIN_FONT_SIZES['subtitle'])
            p.font.color.rgb = RgbColor.from_string('FFFFFF')
            y += 0.9

        # 要点
        if content.get('bullets'):
            for bullet in content['bullets'][:ManusConfig.CONTENT_LIMITS['max_bullets_per_slide']]:
                bullet_box = slide.shapes.add_textbox(
                    Inches(0.5), Inches(y),
                    Inches(12), Inches(0.6)
                )
                tf = bullet_box.text_frame
                p = tf.paragraphs[0]
                p.text = f"• {bullet}"
                p.font.size = Pt(ManusConfig.MIN_FONT_SIZES['body'])
                p.font.color.rgb = RgbColor.from_string('E0E0E0')
                y += 0.7

    def _get_default_outline(self, topic: str) -> List[Dict]:
        """获取默认大纲结构"""
        return [
            {"title": topic, "subtitle": datetime.now().strftime("%Y年%m月%d日"), "type": "cover"},
            {"title": "概述", "bullets": ["背景介绍", "核心内容", "主要目标"], "type": "content"},
            {"title": "详细分析", "bullets": ["数据分析", "关键发现", "趋势洞察"], "type": "content"},
            {"title": "核心洞察", "bullets": ["洞察1", "洞察2", "洞察3"], "type": "content"},
            {"title": "行动建议", "bullets": ["建议1", "建议2", "建议3"], "type": "content"},
            {"title": "总结", "bullets": ["要点回顾", "下一步计划"], "type": "content"},
            {"title": "感谢观看", "subtitle": "Thank You", "type": "summary"}
        ]

    @classmethod
    async def from_dashboard(
        cls,
        gemini_api_key: str,
        claude_api_key: str,
        dashboard_data: Dict[str, Any],
        style: str = "professional"
    ) -> str:
        """从仪表盘数据生成PPT"""
        title = dashboard_data.get('title', '数据分析报告')
        kpis = dashboard_data.get('kpis', [])
        charts = dashboard_data.get('charts', [])
        insights = dashboard_data.get('insights', [])

        slides = []

        # 封面
        slides.append({
            "title": title,
            "subtitle": f"报告日期: {datetime.now().strftime('%Y年%m月%d日')}",
            "type": "cover"
        })

        # KPI页
        if kpis:
            slides.append({
                "title": "关键业绩指标",
                "bullets": [f"{kpi['label']}: {kpi['value']}" for kpi in kpis],
                "type": "content"
            })

        # 图表分析页
        for chart in charts:
            slides.append({
                "title": chart.get('title', ''),
                "subtitle": chart.get('description', ''),
                "bullets": chart.get('highlights', []),
                "type": "content"
            })

        # 洞察页
        if insights:
            slides.append({
                "title": "核心洞察与建议",
                "bullets": insights,
                "type": "content"
            })

        # 结尾
        slides.append({
            "title": "感谢观看",
            "subtitle": "Thank You",
            "content": "由 Manus Clone PPT Generator 生成\n双模型协作: Gemini + Claude",
            "type": "summary"
        })

        config = ManusCloneConfig(
            title=title,
            style=style,
            num_slides=len(slides),
            use_8k=True
        )

        generator = cls(gemini_api_key, claude_api_key, config)
        return await generator.generate(title, "", slides)


# ============================================================================
# CLI入口
# ============================================================================

async def main():
    """主入口"""
    import argparse

    parser = argparse.ArgumentParser(description="🍌 Manus Clone PPT Generator")
    parser.add_argument("-t", "--topic", default="数据分析报告", help="PPT主题")
    parser.add_argument("-s", "--style", default="professional",
                        choices=["professional", "minimal_8k", "dark_tech_8k", "gradient_premium_8k", "corporate_8k"],
                        help="幻灯片风格")
    parser.add_argument("-n", "--num-slides", type=int, default=8, help="幻灯片数量")
    parser.add_argument("-d", "--from-dashboard", action="store_true", help="从仪表盘数据生成")
    parser.add_argument("--no-8k", action="store_true", help="禁用8K分辨率")
    parser.add_argument("--no-quality-check", action="store_true", help="禁用质量检查")

    args = parser.parse_args()

    # 获取API密钥
    gemini_key = os.getenv("GEMINI_API_KEY")
    claude_key = os.getenv("CLAUDE_API_KEY")

    if not gemini_key or not claude_key:
        console.print("[red]错误: 请设置 GEMINI_API_KEY 和 CLAUDE_API_KEY 环境变量[/red]")
        return

    if args.from_dashboard:
        # 使用格力仪表盘数据
        dashboard_data = {
            'title': '格力抖音官方旗舰店 - 数据分析报告',
            'kpis': [
                {'label': '总销售额', 'value': '¥6.55亿'},
                {'label': '总销量', 'value': '24.8万台'},
                {'label': '整体客单价', 'value': '¥2,642'},
                {'label': '业务周期', 'value': '40个月'}
            ],
            'charts': [
                {
                    'title': '年度销售趋势分析',
                    'description': '2022-2025年销售额与销量变化',
                    'highlights': ['2023年同比增长546%', '2024年达峰值3.37亿', '销量持续增长']
                },
                {
                    'title': '渠道分布分析',
                    'description': '各渠道销售占比',
                    'highlights': ['直播渠道贡献76%', '商品卡贡献16.5%', '其他渠道7.5%']
                },
                {
                    'title': '季节性分析',
                    'description': '月度销售波动规律',
                    'highlights': ['4-6月为旺季', '6月峰值4656万', '1-2月和8月淡季']
                }
            ],
            'insights': [
                '直播电商是核心增长引擎',
                '挂机空调是主力产品',
                '需提前布局旺季备货',
                '建议加强商品卡渠道'
            ]
        }

        output = await ManusClonePPTGenerator.from_dashboard(
            gemini_key, claude_key, dashboard_data, args.style
        )
    else:
        config = ManusCloneConfig(
            title=args.topic,
            style=args.style,
            num_slides=args.num_slides,
            use_8k=not args.no_8k,
            enable_quality_check=not args.no_quality_check
        )

        generator = ManusClonePPTGenerator(gemini_key, claude_key, config)
        output = await generator.generate(args.topic)

    console.print(f"\n[bold green]🎉 完成![/bold green] 文件: {output}")


if __name__ == "__main__":
    asyncio.run(main())
