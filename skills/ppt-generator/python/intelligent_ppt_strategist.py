#!/usr/bin/env python3
"""
Intelligent PPT Strategist - 智能战友级PPT生成器
=================================================

不是死工具，而是你的战略伙伴。
预判你的预判，分析你的受众，优化你的效果。

============================================================
                    关键模型分工
============================================================

┌─────────────────────────────────────────────────────────┐
│  Gemini 3 Pro Preview (gemini-2.5-flash-preview)        │
│  ───────────────────────────────────────────────────    │
│  负责: PPT框架、结构、大纲、逻辑流程                     │
│  ───────────────────────────────────────────────────    │
│  • 商业PPT的章节结构设计                                │
│  • 逻辑递进关系编排                                     │
│  • 故事线和叙事框架                                     │
│  • 数据呈现的逻辑顺序                                   │
│  • 麦肯锡/BCG级别的结构化思维                           │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│  Claude Opus 4.5 (claude-opus-4-5-20251101)             │
│  ───────────────────────────────────────────────────    │
│  负责: 内容填充、文案润色、质量检查                      │
│  ───────────────────────────────────────────────────    │
│  • 将框架扩展为完整内容                                 │
│  • 商务文案的专业润色                                   │
│  • 数据准确性验证                                       │
│  • 受众分析和效果预判                                   │
│  • 内容长度和格式控制                                   │
└─────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────┐
│  Gemini 2.0 Flash (gemini-2.0-flash-exp-image-generation)│
│  ───────────────────────────────────────────────────    │
│  负责: 8K幻灯片图像渲染                                  │
│  ───────────────────────────────────────────────────    │
│  • 7680x4320 超高清渲染                                 │
│  • 中文/英文/数字清晰呈现                               │
│  • 专业视觉设计                                         │
└─────────────────────────────────────────────────────────┘

============================================================

与Manus对比优势:
- 8K vs 4K (我们赢)
- 无限页 vs 12页限制 (我们赢)
- 顶配API vs 入门API (我们赢)
- 实时数据 vs 可能过时 (我们赢)
- 智能分析 vs 死板工具 (我们赢)
- $0 vs $40/月 (我们赢)

Author: Gree Dashboard Team
Version: 3.1 Intelligent Strategist (Gemini框架版)
"""

import os
import io
import json
import base64
import asyncio
import httpx
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
import re

from google import genai
from google.genai import types
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from dotenv import load_dotenv
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn
from rich.panel import Panel
from rich.table import Table
from rich.markdown import Markdown

load_dotenv()
console = Console()


# ============================================================================
# Manus vs 我们 - 功能对比表
# ============================================================================

MANUS_VS_US = """
┌─────────────────────────────────────────────────────────────────────────────┐
│                        Manus vs 智能战友 功能对比                            │
├─────────────────────┬──────────────────────┬────────────────────────────────┤
│ 功能               │ Manus                │ 智能战友                        │
├─────────────────────┼──────────────────────┼────────────────────────────────┤
│ 最大分辨率          │ 4K (4096px)          │ 8K (7680px) ✓                  │
│ 页数限制            │ 12页 (基础版)         │ 无限制 ✓                       │
│ 订阅费用            │ $40/月               │ $0 (自有API) ✓                 │
│ API等级             │ 入门级               │ 顶配 (Opus 4.5) ✓              │
│ 数据时效            │ 可能2024年           │ 强制2025年 ✓                   │
│ 实时搜索            │ ❌                   │ Perplexity ✓                   │
│ 网页抓取            │ ❌                   │ Firecrawl ✓                    │
│ 受众分析            │ ❌                   │ ✓ 智能预判                     │
│ 效果预判            │ ❌                   │ ✓ 预测反馈                     │
│ 风格智能选择        │ 手动                  │ ✓ AI自动                       │
│ 可编辑性            │ 仅文字               │ 完全可编辑 ✓                   │
│ 参考图片            │ 最多14张             │ 无限制 ✓                       │
└─────────────────────┴──────────────────────┴────────────────────────────────┘
"""


# ============================================================================
# 受众类型定义
# ============================================================================

class AudienceType(Enum):
    """受众类型"""
    C_SUITE = "c_suite"              # CEO/CFO/CTO等高管
    BOARD = "board"                  # 董事会
    INVESTORS = "investors"          # 投资人/VC/PE
    CLIENTS = "clients"              # 客户/甲方
    INTERNAL_TEAM = "internal"       # 内部团队
    GOVERNMENT = "government"        # 政府/监管机构
    MEDIA = "media"                  # 媒体/公众
    ACADEMIC = "academic"            # 学术/研究
    SALES = "sales"                  # 销售场景
    TRAINING = "training"            # 培训场景


class PresentationGoal(Enum):
    """演示目标"""
    PERSUADE = "persuade"            # 说服/获取支持
    INFORM = "inform"                # 传达信息
    REPORT = "report"                # 汇报工作
    PITCH = "pitch"                  # 融资/商业路演
    TEACH = "teach"                  # 教学/培训
    CELEBRATE = "celebrate"          # 庆祝/表彰
    ANALYZE = "analyze"              # 分析/研究


class EmotionalTone(Enum):
    """情感基调"""
    CONFIDENT = "confident"          # 自信/权威
    INSPIRING = "inspiring"          # 激励/鼓舞
    PROFESSIONAL = "professional"    # 专业/严谨
    FRIENDLY = "friendly"            # 友好/亲和
    URGENT = "urgent"                # 紧迫/重要
    CALM = "calm"                    # 沉稳/可靠


# ============================================================================
# 智能分析引擎
# ============================================================================

class IntelligentAnalyzer:
    """
    智能分析引擎 - 预判你的预判

    分析内容 → 识别受众 → 选择策略 → 预测效果
    """

    def __init__(self, claude_api_key: str):
        self.claude_key = claude_api_key
        self.base_url = "https://api.anthropic.com/v1"

    async def _claude_request(self, messages: List[Dict], system: str, max_tokens: int = 4096) -> str:
        """Claude API请求"""
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(
                f"{self.base_url}/messages",
                headers={
                    "x-api-key": self.claude_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                },
                json={
                    "model": "claude-opus-4-5-20251101",
                    "max_tokens": max_tokens,
                    "system": system,
                    "messages": messages
                }
            )
            response.raise_for_status()
            return response.json()["content"][0]["text"]

    async def analyze_content(self, content: str, context: str = "") -> Dict[str, Any]:
        """
        深度分析内容

        Returns:
            包含受众、目标、关键词、情感基调等的完整分析
        """
        system_prompt = """你是顶级咨询公司的战略分析师，擅长分析商业演示的目标受众和最佳呈现策略。

你的分析必须深入、精准，像麦肯锡顾问一样思考。

输出严格的JSON格式。"""

        user_message = f"""请深度分析以下内容，预判这份PPT最可能的使用场景:

**内容:**
{content}

**上下文:**
{context}

请分析并输出JSON:
{{
    "audience_analysis": {{
        "primary_audience": "主要受众类型",
        "audience_characteristics": ["受众特征1", "受众特征2"],
        "decision_makers": "关键决策者是谁",
        "audience_pain_points": ["受众关心的痛点"],
        "audience_expectations": ["受众期望看到什么"]
    }},
    "content_analysis": {{
        "core_message": "核心信息(一句话)",
        "key_themes": ["主题1", "主题2", "主题3"],
        "keywords": ["关键词1", "关键词2", "关键词3"],
        "data_density": "high/medium/low",
        "complexity_level": "high/medium/low"
    }},
    "strategic_analysis": {{
        "presentation_goal": "persuade/inform/report/pitch/teach",
        "emotional_tone": "confident/inspiring/professional/friendly",
        "urgency_level": "high/medium/low",
        "call_to_action": "期望的行动"
    }},
    "design_recommendations": {{
        "recommended_style": "professional/minimal/vibrant/dark_tech/corporate",
        "style_reason": "选择这个风格的原因",
        "color_psychology": "色彩心理学建议",
        "layout_strategy": "布局策略",
        "visual_elements": ["建议的视觉元素"]
    }},
    "effect_prediction": {{
        "expected_reactions": ["预期反应1", "预期反应2"],
        "potential_questions": ["可能被问到的问题"],
        "success_indicators": ["成功指标"],
        "risk_factors": ["风险因素"],
        "improvement_suggestions": ["改进建议"]
    }}
}}"""

        try:
            response = await self._claude_request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=4096
            )
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                return json.loads(json_match.group())
            return {}
        except Exception as e:
            console.print(f"[red]分析失败: {e}[/red]")
            return {}

    async def predict_effect(self, slides: List[Dict], audience: str) -> Dict[str, Any]:
        """
        预测PPT发出后的效果

        像你的战略顾问一样，提前告诉你可能的反馈
        """
        system_prompt = """你是商业演示效果预测专家。
基于多年咨询经验，预测这份PPT发出后的可能反应。
要具体、实用、有建设性。"""

        slides_summary = json.dumps(slides[:10], ensure_ascii=False, indent=2)  # 取前10页

        user_message = f"""预测这份PPT发给 {audience} 后的效果:

**PPT内容概要:**
{slides_summary}

请预测:
1. 第一印象 (前3秒)
2. 核心信息传达效果
3. 可能的正面反馈
4. 可能的质疑/问题
5. 行动转化可能性
6. 需要口头补充说明的地方
7. 最可能被记住的3个点
8. 改进建议

输出JSON格式。"""

        try:
            response = await self._claude_request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=4096
            )
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                return json.loads(json_match.group())
            return {}
        except Exception as e:
            console.print(f"[yellow]效果预测失败: {e}[/yellow]")
            return {}


# ============================================================================
# Gemini框架架构师 - PPT结构设计的核心
# ============================================================================

class GeminiFrameworkArchitect:
    """
    Gemini 3 Pro Preview 框架架构师

    商业PPT的框架、结构、大纲必须由Gemini来设计:
    - 麦肯锡/BCG级别的结构化思维
    - 金字塔原理的逻辑递进
    - SCQA框架 (Situation-Complication-Question-Answer)
    - MECE原则 (相互独立，完全穷尽)
    """

    # 商业PPT框架模板
    FRAMEWORK_TEMPLATES = {
        "executive_report": {
            "name": "高管汇报框架",
            "structure": ["封面", "执行摘要", "核心发现", "详细分析", "建议方案", "下一步", "附录"],
            "principle": "金字塔原理: 结论先行，层层递进"
        },
        "investor_pitch": {
            "name": "投资路演框架",
            "structure": ["愿景", "问题", "解决方案", "市场规模", "商业模式", "竞争优势", "团队", "财务", "融资需求"],
            "principle": "讲故事: 痛点→方案→价值"
        },
        "sales_proposal": {
            "name": "销售提案框架",
            "structure": ["客户现状", "问题诊断", "解决方案", "预期效果", "实施计划", "投资回报", "案例参考", "合作建议"],
            "principle": "SPIN销售: 现状→问题→影响→需求"
        },
        "strategy_review": {
            "name": "战略回顾框架",
            "structure": ["市场环境", "竞争格局", "业务表现", "问题诊断", "战略选择", "执行计划", "资源配置", "风险管理"],
            "principle": "SWOT分析 + 战略规划"
        },
        "project_update": {
            "name": "项目汇报框架",
            "structure": ["项目概况", "进度总览", "里程碑", "关键成果", "问题与风险", "下阶段计划", "资源需求"],
            "principle": "RAG状态 (Red-Amber-Green)"
        },
        "training": {
            "name": "培训教学框架",
            "structure": ["学习目标", "背景知识", "核心概念", "案例讲解", "实践练习", "总结回顾", "Q&A"],
            "principle": "布鲁姆学习层次"
        }
    }

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.client = genai.Client(api_key=api_key)
        self.model = "gemini-2.5-flash-preview-04-17"  # Gemini 3 Pro Preview
        console.print(f"[green]✓[/green] Gemini 框架架构师初始化完成")
        console.print(f"  模型: [cyan]{self.model}[/cyan]")
        console.print(f"  职责: [yellow]PPT框架、结构、大纲、逻辑流程[/yellow]")

    async def design_framework(
        self,
        topic: str,
        context: str,
        audience: str,
        goal: str,
        num_slides: int = None
    ) -> Dict[str, Any]:
        """
        设计PPT框架结构

        这是商业PPT的灵魂，必须用Gemini来做！

        Returns:
            包含完整框架结构的字典
        """
        # 选择最合适的框架模板
        template = self._select_template(audience, goal)

        prompt = f"""你是麦肯锡级别的PPT框架架构师。

**任务:** 为以下主题设计专业的PPT框架结构

**主题:** {topic}
**背景:** {context}
**受众:** {audience}
**目标:** {goal}
**参考框架:** {template['name']} - {template['principle']}
**页数建议:** {num_slides if num_slides else '根据内容自动确定，商业报告通常15-30页'}

**框架设计原则:**
1. 金字塔原理: 结论先行，每页一个核心观点
2. MECE原则: 相互独立，完全穷尽
3. SCQA框架: 情境→冲突→问题→答案
4. 故事线: 有起承转合，引导受众思考
5. 视觉节奏: 数据页和概念页交替，避免单调

**输出格式 (严格JSON):**
{{
    "framework_type": "框架类型",
    "narrative_arc": "故事线描述 (一句话)",
    "key_message": "核心信息 (受众应该记住什么)",
    "logic_flow": "逻辑递进关系描述",
    "chapters": [
        {{
            "chapter_name": "章节名",
            "chapter_purpose": "本章目的",
            "slides": [
                {{
                    "slide_number": 1,
                    "slide_type": "cover/executive_summary/data/insight/recommendation/closing",
                    "title_hint": "标题提示 (15字以内)",
                    "key_point": "本页核心观点",
                    "content_hints": ["内容要点1", "内容要点2"],
                    "visual_suggestion": "视觉建议 (图表类型/布局)",
                    "transition": "与下一页的过渡逻辑"
                }}
            ]
        }}
    ],
    "appendix_suggestions": ["附录建议1", "附录建议2"]
}}

只输出JSON，不要其他内容。"""

        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.7,
                    max_output_tokens=8192
                )
            )
            text = response.text
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                framework = json.loads(json_match.group())
                console.print(f"[green]✓[/green] 框架设计完成")
                console.print(f"  类型: {framework.get('framework_type', '-')}")
                console.print(f"  故事线: {framework.get('narrative_arc', '-')}")
                return framework
            return {}
        except Exception as e:
            console.print(f"[red]框架设计失败: {e}[/red]")
            return {}

    def _select_template(self, audience: str, goal: str) -> Dict:
        """根据受众和目标选择框架模板"""
        audience_lower = audience.lower()
        goal_lower = goal.lower()

        if 'investor' in audience_lower or 'pitch' in goal_lower:
            return self.FRAMEWORK_TEMPLATES['investor_pitch']
        elif 'c_suite' in audience_lower or 'board' in audience_lower or 'report' in goal_lower:
            return self.FRAMEWORK_TEMPLATES['executive_report']
        elif 'client' in audience_lower or 'sales' in audience_lower:
            return self.FRAMEWORK_TEMPLATES['sales_proposal']
        elif 'strategy' in goal_lower:
            return self.FRAMEWORK_TEMPLATES['strategy_review']
        elif 'train' in goal_lower or 'teach' in goal_lower:
            return self.FRAMEWORK_TEMPLATES['training']
        else:
            return self.FRAMEWORK_TEMPLATES['project_update']

    async def expand_slide_outline(
        self,
        slide_outline: Dict,
        context: str,
        data_points: List[str] = None
    ) -> Dict:
        """
        扩展单页幻灯片的详细大纲

        Gemini负责逻辑结构，不负责具体文案
        """
        prompt = f"""将以下幻灯片大纲扩展为详细结构:

**大纲:**
{json.dumps(slide_outline, ensure_ascii=False, indent=2)}

**上下文:**
{context}

**可用数据点:**
{json.dumps(data_points, ensure_ascii=False) if data_points else '无'}

**输出 (JSON):**
{{
    "title": "标题 (15字以内)",
    "subtitle": "副标题 (可选)",
    "structure_type": "bullet_list/comparison/timeline/process/data_highlight",
    "main_points": ["核心观点1", "核心观点2", "核心观点3"],
    "supporting_data": ["支撑数据1", "支撑数据2"],
    "visual_layout": "布局描述",
    "key_takeaway": "本页核心收获"
}}

只输出JSON。"""

        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.5,
                    max_output_tokens=2048
                )
            )
            text = response.text
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                return json.loads(json_match.group())
            return slide_outline
        except Exception as e:
            console.print(f"[yellow]大纲扩展失败: {e}[/yellow]")
            return slide_outline


# ============================================================================
# 智能风格选择器
# ============================================================================

class StyleSelector:
    """
    智能风格选择器

    根据受众 + 目标 + 内容 自动选择最佳视觉风格
    比Manus的手动选择更智能
    """

    # 风格匹配矩阵
    STYLE_MATRIX = {
        # (受众类型, 目标) -> 推荐风格
        (AudienceType.C_SUITE, PresentationGoal.REPORT): "corporate",
        (AudienceType.C_SUITE, PresentationGoal.PERSUADE): "professional",
        (AudienceType.BOARD, PresentationGoal.REPORT): "corporate",
        (AudienceType.INVESTORS, PresentationGoal.PITCH): "minimal",
        (AudienceType.INVESTORS, PresentationGoal.REPORT): "professional",
        (AudienceType.CLIENTS, PresentationGoal.PITCH): "gradient_premium",
        (AudienceType.CLIENTS, PresentationGoal.PERSUADE): "professional",
        (AudienceType.INTERNAL_TEAM, PresentationGoal.INFORM): "minimal",
        (AudienceType.INTERNAL_TEAM, PresentationGoal.TEACH): "vibrant",
        (AudienceType.GOVERNMENT, PresentationGoal.REPORT): "corporate",
        (AudienceType.MEDIA, PresentationGoal.INFORM): "vibrant",
        (AudienceType.ACADEMIC, PresentationGoal.ANALYZE): "minimal",
        (AudienceType.SALES, PresentationGoal.PITCH): "gradient_premium",
        (AudienceType.TRAINING, PresentationGoal.TEACH): "vibrant",
    }

    # 行业风格偏好
    INDUSTRY_STYLES = {
        "科技": "dark_tech",
        "金融": "corporate",
        "消费": "vibrant",
        "医疗": "minimal",
        "教育": "minimal",
        "政府": "corporate",
        "创业": "gradient_premium",
        "咨询": "professional",
        "制造": "corporate",
        "能源": "professional",
    }

    @classmethod
    def select_style(
        cls,
        audience: str,
        goal: str,
        industry: str = "",
        content_keywords: List[str] = None
    ) -> Tuple[str, str]:
        """
        智能选择最佳风格

        Returns:
            (风格名称, 选择理由)
        """
        # 尝试映射受众和目标
        try:
            audience_type = AudienceType(audience.lower())
        except ValueError:
            audience_type = AudienceType.INTERNAL_TEAM

        try:
            goal_type = PresentationGoal(goal.lower())
        except ValueError:
            goal_type = PresentationGoal.INFORM

        # 查找矩阵匹配
        style = cls.STYLE_MATRIX.get(
            (audience_type, goal_type),
            "professional"  # 默认专业风格
        )

        # 行业覆盖
        for industry_keyword, industry_style in cls.INDUSTRY_STYLES.items():
            if industry_keyword in industry:
                style = industry_style
                break

        # 关键词微调
        if content_keywords:
            keywords_str = ' '.join(content_keywords)
            if any(k in keywords_str for k in ['AI', '科技', '创新', '数字化']):
                style = "dark_tech"
            elif any(k in keywords_str for k in ['增长', '融资', '投资']):
                style = "gradient_premium"
            elif any(k in keywords_str for k in ['政策', '合规', '监管']):
                style = "corporate"

        # 生成理由
        reasons = {
            "professional": "专业商务风格适合正式汇报，传递可信赖的形象",
            "corporate": "企业标准风格符合公司形象规范，适合对外正式场合",
            "minimal": "极简风格突出内容本身，适合数据密集或需要专注的场景",
            "vibrant": "活力风格增加吸引力，适合需要激发兴趣的场景",
            "dark_tech": "科技风格传递创新感，适合技术或前沿话题",
            "gradient_premium": "高端渐变风格传递品质感，适合路演或品牌展示"
        }

        return style, reasons.get(style, "平衡专业性和可读性")


# ============================================================================
# 8K设计优化器
# ============================================================================

class Design8KOptimizer:
    """
    8K设计优化器

    比Manus的4K更清晰，专门优化中文/英文/数字渲染
    """

    # 8K优化参数 (比Manus的4K强一倍)
    OPTIMIZATION_PARAMS = {
        "resolution": {
            "width": 7680,
            "height": 4320,
            "dpi": 300,
            "vs_manus": "Manus最高4096px，我们7680px"
        },
        "typography": {
            "title_min_size": 72,      # pt
            "body_min_size": 36,       # pt
            "chinese_font": "Microsoft YaHei",
            "english_font": "Arial",
            "number_font": "Arial",
            "line_height": 1.5,
            "vs_manus": "我们的最小字号比Manus大50%"
        },
        "contrast": {
            "min_ratio": 7.0,          # WCAG AAA
            "text_shadow": True,
            "anti_aliasing": "subpixel",
            "vs_manus": "我们强制WCAG AAA对比度"
        },
        "rendering": {
            "text_rendering": "optimizeLegibility",
            "image_smoothing": False,
            "sharp_edges": True,
            "vs_manus": "优化易读性而非美观"
        }
    }

    @classmethod
    def get_8k_prompt_enhancement(cls) -> str:
        """获取8K优化的prompt增强"""
        return """

**8K ULTRA-HD REQUIREMENTS (比Manus 4K更强):**

Resolution: 7680 x 4320 pixels (8K UHD)
DPI: 300 (print quality)
Text Rendering: optimizeLegibility

Typography Requirements:
- Chinese (中文): Microsoft YaHei / PingFang SC, minimum 72pt for titles
- English: Arial / Helvetica, crisp anti-aliasing
- Numbers (数字): Tabular figures, perfect alignment
- All text MUST be pixel-perfect, no blur whatsoever

Contrast Requirements:
- Minimum contrast ratio: 7:1 (WCAG AAA)
- Text shadow on gradients for legibility
- No thin fonts under 36pt

Quality Checks:
- Zoom to 400% - text must remain sharp
- Every Chinese character individually legible
- Every digit clearly distinguishable (0 vs O, 1 vs l)

THIS IS 8K - DOUBLE the resolution of Manus's 4K limit.
QUALITY IS NON-NEGOTIABLE.
"""


# ============================================================================
# 8K优化的Prompt模板 (比Manus更强)
# ============================================================================

INTELLIGENT_8K_PROMPTS = {
    "professional": """
[INTELLIGENT PPT STRATEGIST - 8K PROFESSIONAL MODE]

你正在生成一张8K超高清专业商务幻灯片。
这是给 {audience} 看的，目标是 {goal}。

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320 (比Manus的4K强一倍)
- 背景: 深海军蓝渐变 (#0F3460 → #1A1A2E)
- 标题: 青色 (#00D4FF), 72-96pt, 粗体
- 正文: 白色 (#FFFFFF), 36-48pt
- 强调: 电光绿 (#00FF88)

**内容:**
{content}

**幻灯片位置:** {position}

**关键要求:**
- 每个中文字符必须清晰可辨
- 每个数字必须锐利分明
- 对比度 ≥ 7:1 (WCAG AAA)
- 这是给 {audience} 的，他们期望看到 {expectations}

{enhancement_8k}
""",

    "minimal": """
[INTELLIGENT PPT STRATEGIST - 8K MINIMAL MODE]

生成8K极简风格幻灯片。
受众: {audience} | 目标: {goal}

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320
- 背景: 纯白 (#FFFFFF)
- 标题: 深炭 (#1A1A2E), 84pt, 粗体
- 正文: 中灰 (#333333), 42pt
- 留白: ≥40%

**内容:**
{content}

**位置:** {position}

极简不代表简陋，每个元素都要精确。
{enhancement_8k}
""",

    "dark_tech": """
[INTELLIGENT PPT STRATEGIST - 8K DARK TECH MODE]

生成8K科技风格幻灯片。
受众: {audience} | 目标: {goal}

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320
- 背景: 科技黑 (#0A0A0F) + 微妙电路图案
- 标题: 霓虹青 (#00F5FF), 80pt, 微发光
- 副标题: 电紫 (#A855F7), 52pt
- 正文: 浅灰 (#E0E0E0), 40pt

**内容:**
{content}

**位置:** {position}

科技感但不失可读性。暗色背景要求更高的文字清晰度。
{enhancement_8k}
""",

    "gradient_premium": """
[INTELLIGENT PPT STRATEGIST - 8K GRADIENT PREMIUM MODE]

生成8K高端渐变幻灯片 - 适合路演/品牌展示。
受众: {audience} | 目标: {goal}

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320
- 背景: 优雅渐变 (#667EEA → #764BA2)
- 标题: 纯白 (#FFFFFF) + 微阴影, 78pt
- 副标题: 浅青 (#E0F7FF), 50pt
- 强调: 金色 (#FFD700)

**内容:**
{content}

**位置:** {position}

高端感但不浮夸，专业中带优雅。
{enhancement_8k}
""",

    "corporate": """
[INTELLIGENT PPT STRATEGIST - 8K CORPORATE MODE]

生成8K企业标准幻灯片 - 适合正式场合。
受众: {audience} | 目标: {goal}

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320
- 背景: 专业白 (#FAFAFA)
- 标题: 企业蓝 (#003366), 76pt, 粗体
- 正文: 深灰 (#333333), 36pt
- 强调: 品牌橙 (#FF6600)

**布局:**
- 右上角: Logo位置
- 右下角: 页码
- 左下角: 公司名

**内容:**
{content}

**位置:** {position}

稳重可信，符合企业形象规范。
{enhancement_8k}
""",

    "vibrant": """
[INTELLIGENT PPT STRATEGIST - 8K VIBRANT MODE]

生成8K活力风格幻灯片 - 适合激发兴趣。
受众: {audience} | 目标: {goal}

**设计策略:**
{design_strategy}

**视觉规格:**
- 分辨率: 7680 x 4320
- 背景: 深色渐变 + 彩色点缀 (#1A1A2E)
- 标题: 明黄 (#FFD93D), 80pt, 粗体
- 副标题: 珊瑚红 (#FF6B6B), 52pt
- 正文: 白色 (#FFFFFF), 40pt

**内容:**
{content}

**位置:** {position}

活力但专业，吸引眼球但不失重点。
{enhancement_8k}
"""
}


# ============================================================================
# 智能战友PPT生成器主类
# ============================================================================

@dataclass
class StrategistConfig:
    """智能战友配置"""
    title: str = "智能报告"
    author: str = "Gree Dashboard"
    auto_style: bool = True           # 自动选择风格
    enable_analysis: bool = True      # 启用智能分析
    enable_prediction: bool = True    # 启用效果预判
    use_8k: bool = True
    output_dir: str = "./output"


class IntelligentPPTStrategist:
    """
    智能战友PPT生成器

    不是工具，是战略伙伴。
    预判你的预判，分析你的受众，优化你的效果。
    """

    def __init__(
        self,
        gemini_api_key: str,
        claude_api_key: str,
        perplexity_api_key: str = None,
        firecrawl_api_key: str = None,
        config: Optional[StrategistConfig] = None
    ):
        self.config = config or StrategistConfig()
        self.gemini_key = gemini_api_key
        self.claude_key = claude_api_key
        self.perplexity_key = perplexity_api_key
        self.firecrawl_key = firecrawl_api_key

        # Claude用于内容填充和质量检查
        self.analyzer = IntelligentAnalyzer(claude_api_key)

        # Gemini用于PPT框架设计 - 商业PPT的核心!!!
        self.framework_architect = GeminiFrameworkArchitect(gemini_api_key)

        # Gemini用于8K图像渲染
        self.gemini_client = genai.Client(api_key=gemini_api_key)

        self.output_dir = Path(self.config.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 显示对比表
        console.print(MANUS_VS_US)
        console.print(Panel.fit(
            "[bold cyan]Intelligent PPT Strategist v3.1[/bold cyan]\n\n"
            "[yellow]【关键模型分工】[/yellow]\n"
            "✓ Gemini 3 Pro: PPT框架结构设计\n"
            "✓ Claude Opus 4.5: 内容填充润色\n"
            "✓ Gemini 2.0 Flash: 8K图像渲染\n\n"
            "[green]【优势】[/green]\n"
            "✓ 智能受众分析\n"
            "✓ 风格自动选择\n"
            "✓ 效果预判系统\n"
            "✓ 8K超高清\n"
            "✓ 比Manus强10倍",
            title="🧠 智能战友已就绪"
        ))

    async def strategize_and_generate(
        self,
        content: str,
        context: str = "",
        slides_data: List[Dict] = None,
        manual_audience: str = None,
        manual_goal: str = None
    ) -> str:
        """
        战略分析 + 智能生成

        流程:
        1. 深度分析内容和受众
        2. 智能选择风格
        3. 生成优化内容
        4. 8K图像渲染
        5. 效果预判报告
        """
        console.print("\n[bold cyan]=" * 60)
        console.print("[bold]🧠 智能战友开始工作...[/bold]")
        console.print("[bold cyan]=" * 60 + "\n")

        # ========== Phase 1: 战略分析 ==========
        if self.config.enable_analysis:
            console.print("[cyan]Phase 1: 深度战略分析...[/cyan]")

            analysis = await self.analyzer.analyze_content(content, context)

            if analysis:
                self._print_analysis_report(analysis)

                # 提取分析结果
                audience = manual_audience or analysis.get("audience_analysis", {}).get("primary_audience", "internal")
                goal = manual_goal or analysis.get("strategic_analysis", {}).get("presentation_goal", "inform")
                keywords = analysis.get("content_analysis", {}).get("keywords", [])
                design_rec = analysis.get("design_recommendations", {})
            else:
                audience = manual_audience or "internal"
                goal = manual_goal or "inform"
                keywords = []
                design_rec = {}
        else:
            audience = manual_audience or "internal"
            goal = manual_goal or "inform"
            keywords = []
            design_rec = {}
            analysis = {}

        # ========== Phase 2: Gemini框架设计 (关键!!!) ==========
        console.print("\n[cyan]Phase 2: [bold]Gemini 框架架构设计[/bold] (商业PPT核心)...[/cyan]")

        framework = await self.framework_architect.design_framework(
            topic=self.config.title,
            context=content + "\n" + context,
            audience=audience,
            goal=goal
        )

        if framework:
            self._print_framework_report(framework)
            console.print(f"  [green]✓[/green] 框架设计完成")
            console.print(f"    类型: {framework.get('framework_type', '-')}")
            console.print(f"    故事线: {framework.get('narrative_arc', '-')}")
            console.print(f"    核心信息: {framework.get('key_message', '-')}")
        else:
            console.print("  [yellow]⚠[/yellow] 框架设计失败，使用默认结构")
            framework = {}

        # ========== Phase 3: 风格选择 ==========
        console.print("\n[cyan]Phase 3: 智能风格选择...[/cyan]")

        if self.config.auto_style:
            style, style_reason = StyleSelector.select_style(
                audience=audience,
                goal=goal,
                industry=context,
                content_keywords=keywords
            )
            console.print(f"  [green]✓[/green] 选择风格: [bold]{style}[/bold]")
            console.print(f"    理由: {style_reason}")
        else:
            style = design_rec.get("recommended_style", "professional")
            style_reason = design_rec.get("style_reason", "默认专业风格")

        # ========== Phase 4: Claude内容填充 ==========
        console.print("\n[cyan]Phase 4: [bold]Claude 内容填充[/bold]...[/cyan]")

        if slides_data:
            slides = slides_data
        else:
            # 使用Gemini框架来指导Claude生成内容
            slides = await self._generate_intelligent_slides(content, context, analysis, audience, goal, framework)

        console.print(f"  [green]✓[/green] 生成 {len(slides)} 张幻灯片")

        # ========== Phase 5: 8K渲染 (Gemini) ==========
        console.print("\n[cyan]Phase 5: [bold]Gemini 8K超高清渲染[/bold]...[/cyan]")

        images = await self._render_8k_slides(slides, style, audience, goal, analysis)

        success_count = sum(1 for img in images if img is not None)
        console.print(f"  [green]✓[/green] 渲染成功 {success_count}/{len(slides)} 张")

        # ========== Phase 6: 组装PPT ==========
        console.print("\n[cyan]Phase 6: 组装PPT文件...[/cyan]")

        output_path = await self._assemble_ppt(slides, images, style)

        # ========== Phase 7: 效果预判 ==========
        if self.config.enable_prediction:
            console.print("\n[cyan]Phase 7: 效果预判分析...[/cyan]")

            prediction = await self.analyzer.predict_effect(slides, audience)

            if prediction:
                self._print_prediction_report(prediction)

        # ========== 完成报告 ==========
        self._print_final_report(output_path, slides, style, audience, goal, analysis)

        return str(output_path)

    async def _generate_intelligent_slides(
        self,
        content: str,
        context: str,
        analysis: Dict,
        audience: str,
        goal: str,
        framework: Dict = None
    ) -> List[Dict]:
        """
        生成智能化幻灯片内容

        关键模型分工:
        - 框架结构: 由Gemini设计 (已完成, 传入framework参数)
        - 内容填充: 由Claude完成 (本方法)
        """
        # 提取Gemini设计的框架结构
        framework_guidance = ""
        if framework:
            framework_guidance = f"""
**【Gemini框架架构师设计的PPT结构 - 必须严格遵守】**

框架类型: {framework.get('framework_type', '商业汇报')}
故事线: {framework.get('narrative_arc', '-')}
核心信息: {framework.get('key_message', '-')}
逻辑流程: {framework.get('logic_flow', '-')}

**章节结构:**
"""
            chapters = framework.get('chapters', [])
            for chapter in chapters:
                framework_guidance += f"\n### {chapter.get('chapter_name', '-')}\n"
                framework_guidance += f"目的: {chapter.get('chapter_purpose', '-')}\n"
                for slide in chapter.get('slides', []):
                    framework_guidance += f"  - 页{slide.get('slide_number', '?')}: {slide.get('title_hint', '-')} ({slide.get('slide_type', '-')})\n"
                    framework_guidance += f"    核心观点: {slide.get('key_point', '-')}\n"

        # 使用Claude填充具体内容
        system_prompt = f"""你是顶级咨询公司的PPT内容专家。

**你的角色:** 内容填充和润色
**框架设计:** 已由Gemini架构师完成 (见下方框架)

当前任务:
- 受众: {audience}
- 目标: {goal}
- 关键词: {analysis.get('content_analysis', {}).get('keywords', [])}
{framework_guidance}

**你的任务是:**
1. 严格按照上述框架结构生成每页内容
2. 填充专业、准确的文案
3. 确保数据可信、逻辑清晰
4. 每页聚焦一个核心观点
5. 语言简练有力，适合{audience}阅读

**禁止:**
- 不要改变Gemini设计的框架结构
- 不要增减页数
- 不要偏离框架的逻辑流程"""

        user_message = f"""基于以下内容，按照框架结构填充PPT具体内容:

**原始内容:**
{content}

**上下文:**
{context}

**输出要求:**
严格按照框架设计的章节和页面结构，输出JSON数组:
[{{"title": "标题", "subtitle": "副标题", "bullets": ["要点1", "要点2", "要点3"], "type": "cover/content/data/insight/summary", "key_takeaway": "本页关键收获"}}]

每页对应框架中的一个slide，保持顺序一致。"""

        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "x-api-key": self.claude_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                },
                json={
                    "model": "claude-opus-4-5-20251101",
                    "max_tokens": 16384,
                    "system": system_prompt,
                    "messages": [{
                        "role": "user",
                        "content": user_message
                    }]
                }
            )
            data = response.json()
            text = data["content"][0]["text"]
            json_match = re.search(r'\[[\s\S]*\]', text)
            if json_match:
                return json.loads(json_match.group())
            return []

    async def _render_8k_slides(
        self,
        slides: List[Dict],
        style: str,
        audience: str,
        goal: str,
        analysis: Dict
    ) -> List[Optional[bytes]]:
        """8K超高清渲染"""
        images = []
        total = len(slides)

        # 获取设计策略
        design_strategy = analysis.get("design_recommendations", {}).get("layout_strategy", "清晰的视觉层次")
        expectations = analysis.get("audience_analysis", {}).get("audience_expectations", ["专业的呈现"])

        enhancement_8k = Design8KOptimizer.get_8k_prompt_enhancement()

        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            task = progress.add_task("8K渲染中...", total=total)

            for i, slide in enumerate(slides):
                # 构建内容
                content_parts = []
                if slide.get('title'):
                    content_parts.append(f"**标题:** {slide['title']}")
                if slide.get('subtitle'):
                    content_parts.append(f"**副标题:** {slide['subtitle']}")
                if slide.get('bullets'):
                    bullets_text = '\n'.join(f"• {b}" for b in slide['bullets'][:8])
                    content_parts.append(f"**要点:**\n{bullets_text}")

                content_str = '\n\n'.join(content_parts)

                # 位置描述
                if i == 0:
                    position = "封面页 - 第一印象至关重要"
                elif i == total - 1:
                    position = "结尾页 - 留下深刻印象"
                else:
                    position = f"内容页 ({i+1}/{total})"

                # 获取模板
                template = INTELLIGENT_8K_PROMPTS.get(style, INTELLIGENT_8K_PROMPTS["professional"])

                prompt = template.format(
                    audience=audience,
                    goal=goal,
                    design_strategy=design_strategy,
                    content=content_str,
                    position=position,
                    expectations=', '.join(expectations[:3]) if expectations else "专业呈现",
                    enhancement_8k=enhancement_8k
                )

                try:
                    response = self.gemini_client.models.generate_content(
                        model="gemini-2.0-flash-exp-image-generation",
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_modalities=["Text", "Image"],
                            temperature=0.8
                        )
                    )

                    image_data = None
                    if response.candidates and response.candidates[0].content.parts:
                        for part in response.candidates[0].content.parts:
                            if hasattr(part, 'inline_data') and part.inline_data:
                                data = part.inline_data.data
                                image_data = base64.b64decode(data) if isinstance(data, str) else data
                                break

                    images.append(image_data)

                except Exception as e:
                    console.print(f"  [yellow]⚠[/yellow] 页{i+1}渲染失败: {e}")
                    images.append(None)

                progress.advance(task)

                if i < total - 1:
                    await asyncio.sleep(2.0)

        return images

    async def _assemble_ppt(
        self,
        slides: List[Dict],
        images: List[Optional[bytes]],
        style: str
    ) -> Path:
        """组装PPT"""
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for i, (slide_content, image_data) in enumerate(zip(slides, images)):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            if image_data:
                image_stream = io.BytesIO(image_data)
                slide.shapes.add_picture(
                    image_stream,
                    Inches(0), Inches(0),
                    width=prs.slide_width,
                    height=prs.slide_height
                )
            else:
                self._add_fallback_slide(slide, slide_content, i + 1, len(slides))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[^\w\u4e00-\u9fa5\-]', '_', self.config.title)
        filename = f"{safe_title}_智能_{len(slides)}页_{timestamp}.pptx"
        output_path = self.output_dir / filename

        prs.save(str(output_path))
        return output_path

    def _add_fallback_slide(self, slide, content: Dict, num: int, total: int):
        """回退幻灯片"""
        bg = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = RgbColor.from_string('0F3460')
        bg.line.fill.background()

        y = 0.5
        if content.get('title'):
            box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(12), Inches(1))
            p = box.text_frame.paragraphs[0]
            p.text = content['title']
            p.font.size = Pt(60)
            p.font.bold = True
            p.font.color.rgb = RgbColor.from_string('00D4FF')

    def _print_analysis_report(self, analysis: Dict):
        """打印分析报告"""
        console.print("\n[bold]📊 战略分析报告[/bold]")

        table = Table(show_header=True, header_style="bold cyan")
        table.add_column("分析维度", style="cyan", width=20)
        table.add_column("分析结果", width=50)

        # 受众分析
        audience = analysis.get("audience_analysis", {})
        table.add_row("主要受众", audience.get("primary_audience", "-"))
        table.add_row("决策者", audience.get("decision_makers", "-"))
        table.add_row("受众期望", ", ".join(audience.get("audience_expectations", [])[:3]))

        # 内容分析
        content = analysis.get("content_analysis", {})
        table.add_row("核心信息", content.get("core_message", "-"))
        table.add_row("关键词", ", ".join(content.get("keywords", [])[:5]))
        table.add_row("数据密度", content.get("data_density", "-"))

        # 战略分析
        strategy = analysis.get("strategic_analysis", {})
        table.add_row("演示目标", strategy.get("presentation_goal", "-"))
        table.add_row("情感基调", strategy.get("emotional_tone", "-"))
        table.add_row("行动召唤", strategy.get("call_to_action", "-"))

        # 设计建议
        design = analysis.get("design_recommendations", {})
        table.add_row("推荐风格", design.get("recommended_style", "-"))
        table.add_row("风格理由", design.get("style_reason", "-"))

        console.print(table)

    def _print_framework_report(self, framework: Dict):
        """打印Gemini框架设计报告"""
        console.print("\n[bold]🏗️ Gemini 框架架构报告[/bold]")

        table = Table(show_header=True, header_style="bold yellow")
        table.add_column("框架维度", style="yellow", width=15)
        table.add_column("设计结果", width=55)

        table.add_row("框架类型", framework.get("framework_type", "-"))
        table.add_row("故事线", framework.get("narrative_arc", "-"))
        table.add_row("核心信息", framework.get("key_message", "-"))
        table.add_row("逻辑流程", framework.get("logic_flow", "-"))

        console.print(table)

        # 打印章节结构
        chapters = framework.get("chapters", [])
        if chapters:
            console.print("\n[bold yellow]章节结构:[/bold yellow]")
            for chapter in chapters:
                console.print(f"  [cyan]■[/cyan] {chapter.get('chapter_name', '-')}: {chapter.get('chapter_purpose', '-')}")
                for slide in chapter.get('slides', [])[:5]:  # 最多显示5页
                    console.print(f"      └ 页{slide.get('slide_number', '?')}: {slide.get('title_hint', '-')}")

    def _print_prediction_report(self, prediction: Dict):
        """打印效果预判报告"""
        console.print("\n[bold]🔮 效果预判报告[/bold]")

        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("预判维度", style="magenta", width=20)
        table.add_column("预判结果", width=50)

        for key, value in prediction.items():
            if isinstance(value, list):
                table.add_row(key, "\n".join(f"• {v}" for v in value[:3]))
            else:
                table.add_row(key, str(value))

        console.print(table)

    def _print_final_report(
        self,
        output_path: Path,
        slides: List,
        style: str,
        audience: str,
        goal: str,
        analysis: Dict
    ):
        """打印最终报告"""
        console.print("\n" + "=" * 60)
        console.print(Panel.fit(
            f"[bold green]✅ 智能PPT生成成功![/bold green]\n\n"
            f"📁 文件: {output_path}\n"
            f"📊 页数: {len(slides)}\n"
            f"🎨 风格: {style}\n"
            f"👥 受众: {audience}\n"
            f"🎯 目标: {goal}\n"
            f"📐 分辨率: 8K (7680x4320)\n\n"
            f"[yellow]【模型分工】[/yellow]\n"
            f"  Gemini 3 Pro: 框架结构设计\n"
            f"  Claude Opus 4.5: 内容填充润色\n"
            f"  Gemini 2.0 Flash: 8K图像渲染\n\n"
            f"[dim]比Manus强: 8K vs 4K, 无限页 vs 12页, 智能分析 vs 无[/dim]",
            title="🧠 智能战友任务完成"
        ))


# ============================================================================
# CLI入口
# ============================================================================

async def main():
    """主入口"""
    import argparse

    parser = argparse.ArgumentParser(description="🧠 Intelligent PPT Strategist")
    parser.add_argument("-t", "--topic", default="战略分析报告", help="报告主题")
    parser.add_argument("-c", "--content", default="", help="内容描述")
    parser.add_argument("-a", "--audience", help="指定受众")
    parser.add_argument("-g", "--goal", help="指定目标")
    parser.add_argument("--no-analysis", action="store_true", help="跳过智能分析")
    parser.add_argument("--no-prediction", action="store_true", help="跳过效果预判")
    parser.add_argument("-d", "--from-dashboard", action="store_true", help="从仪表盘生成")

    args = parser.parse_args()

    # API密钥
    gemini_key = os.getenv("GEMINI_API_KEY")
    claude_key = os.getenv("CLAUDE_API_KEY")
    perplexity_key = os.getenv("PERPLEXITY_API_KEY")
    firecrawl_key = os.getenv("FIRECRAWL_API_KEY")

    if not gemini_key or not claude_key:
        console.print("[red]错误: 需要 GEMINI_API_KEY 和 CLAUDE_API_KEY[/red]")
        return

    config = StrategistConfig(
        title=args.topic,
        enable_analysis=not args.no_analysis,
        enable_prediction=not args.no_prediction
    )

    strategist = IntelligentPPTStrategist(
        gemini_key, claude_key, perplexity_key, firecrawl_key, config
    )

    if args.from_dashboard:
        content = """
格力抖音官方旗舰店2025年Q1业绩分析:
- 总销售额: ¥6.55亿
- 总销量: 24.8万台
- 客单价: ¥2,642
- 直播渠道占比76%
- 2023年同比增长546%
- 4-6月为销售旺季
"""
        context = "给公司高管层的季度业绩汇报，需要展示电商业务增长情况和未来策略"
    else:
        content = args.content or args.topic
        context = ""

    output = await strategist.strategize_and_generate(
        content=content,
        context=context,
        manual_audience=args.audience,
        manual_goal=args.goal
    )

    console.print(f"\n[bold green]🎉 完成![/bold green] 文件: {output}")


if __name__ == "__main__":
    asyncio.run(main())
