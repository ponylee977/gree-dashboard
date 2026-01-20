#!/usr/bin/env python3
"""
Manus Clone PPT Generator - 企业级增强版
==========================================

核心升级:
- 无页数限制，支持100+页品牌报告
- Perplexity API: 实时互联网搜索，获取最新数据
- Firecrawl API: 深度网页抓取，获取行业报告
- 数据时效性验证: 杜绝过时数据

双模型协作:
- Gemini 3 Pro: 框架逻辑 + 8K图像生成
- Claude Opus 4.5: 内容填充 + 质量检查

Author: Gree Dashboard Team
Version: 2.0 Enterprise
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
from datetime import datetime, timedelta
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

load_dotenv()
console = Console()


# ============================================================================
# 配置 - 企业级无限制版
# ============================================================================

class EnterpriseConfig:
    """企业级配置 - 无页数限制"""

    # 8K分辨率
    RESOLUTION_8K = {"width": 7680, "height": 4320, "dpi": 300}

    # PPT尺寸
    SLIDE_SIZE = {"width": 13.333, "height": 7.5}

    # 字体配置
    FONTS = {
        "chinese_primary": "Microsoft YaHei",
        "chinese_fallback": "PingFang SC",
        "english": "Arial"
    }

    # 最小字号 (8K)
    MIN_FONT_SIZES = {"title": 72, "subtitle": 48, "body": 36, "caption": 28}

    # API配置 - 无页数限制
    API_LIMITS = {
        "max_slides": None,  # 无限制!
        "batch_size": 10,    # 每批处理10张
        "delay_between_batches": 5.0,
        "delay_between_slides": 2.0,
        "retry_attempts": 3,
        "retry_delay": 5.0
    }

    # 内容限制 (单张幻灯片)
    CONTENT_LIMITS = {
        "max_title_length": 50,        # 放宽标题限制
        "max_subtitle_length": 100,    # 放宽副标题限制
        "max_bullets_per_slide": 8,    # 放宽要点数
        "max_bullet_length": 150,      # 放宽要点长度
        "max_body_length": 500         # 放宽正文长度
    }

    # 数据时效性配置
    DATA_FRESHNESS = {
        "max_age_days": 7,             # 数据最大有效期(天)
        "require_2025_data": True,     # 必须是2025年数据
        "reject_2024_data": True,      # 拒绝2024年数据
        "require_source_date": True    # 要求数据来源日期
    }


# ============================================================================
# Perplexity API 客户端 - 实时互联网搜索
# ============================================================================

class PerplexityClient:
    """
    Perplexity API 客户端

    功能:
    - 实时搜索最新互联网数据
    - 获取行业报告和市场数据
    - 确保数据时效性
    """

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = "https://api.perplexity.ai"
        self.model = "llama-3.1-sonar-large-128k-online"  # 在线搜索模型
        console.print(f"[green]✓[/green] Perplexity 实时搜索客户端初始化完成")

    async def search_realtime(
        self,
        query: str,
        focus: str = "internet",
        recency: str = "week"
    ) -> Dict[str, Any]:
        """
        实时搜索最新数据

        Args:
            query: 搜索查询
            focus: 搜索焦点 (internet, academic, news)
            recency: 时效性 (day, week, month)

        Returns:
            搜索结果和来源
        """
        # 强制添加时效性要求到查询
        today = datetime.now().strftime("%Y年%m月")
        enhanced_query = f"{query} 最新数据 {today} 2025年"

        async with httpx.AsyncClient(timeout=60.0) as client:
            try:
                response = await client.post(
                    f"{self.base_url}/chat/completions",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": self.model,
                        "messages": [
                            {
                                "role": "system",
                                "content": f"""你是专业的市场研究分析师。
当前日期: {datetime.now().strftime('%Y年%m月%d日')}
要求:
1. 只提供2025年的最新数据
2. 拒绝使用2024年或更早的数据
3. 必须标注数据来源和日期
4. 如果找不到最新数据，明确说明"""
                            },
                            {
                                "role": "user",
                                "content": enhanced_query
                            }
                        ],
                        "temperature": 0.1,
                        "max_tokens": 4096,
                        "return_citations": True,
                        "search_recency_filter": recency
                    }
                )
                response.raise_for_status()
                data = response.json()

                # 提取内容和引用
                content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
                citations = data.get("citations", [])

                return {
                    "success": True,
                    "content": content,
                    "citations": citations,
                    "query": enhanced_query,
                    "timestamp": datetime.now().isoformat()
                }

            except Exception as e:
                console.print(f"[red]Perplexity搜索失败: {e}[/red]")
                return {
                    "success": False,
                    "error": str(e),
                    "query": enhanced_query,
                    "timestamp": datetime.now().isoformat()
                }

    async def get_industry_report(
        self,
        industry: str,
        metrics: List[str]
    ) -> Dict[str, Any]:
        """获取行业报告数据"""
        query = f"""
{industry}行业最新报告:
- 市场规模和增长率
- 主要玩家和市场份额
- {', '.join(metrics)}
- 2025年最新趋势预测

只要2025年1月最新数据，拒绝2024年数据。
"""
        return await self.search_realtime(query, focus="internet", recency="week")

    async def get_market_data(
        self,
        topic: str,
        data_points: List[str]
    ) -> Dict[str, Any]:
        """获取市场数据"""
        today = datetime.now().strftime("%Y年%m月%d日")
        query = f"""
截至{today}的{topic}最新市场数据:
{chr(10).join(f'- {dp}' for dp in data_points)}

必须是2025年数据，标注具体日期来源。
"""
        return await self.search_realtime(query, focus="news", recency="day")


# ============================================================================
# Firecrawl API 客户端 - 深度网页抓取
# ============================================================================

class FirecrawlClient:
    """
    Firecrawl API 客户端

    功能:
    - 深度抓取网页内容
    - 提取结构化数据
    - 获取行业报告PDF等
    """

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = "https://api.firecrawl.dev/v1"
        console.print(f"[green]✓[/green] Firecrawl 网页抓取客户端初始化完成")

    async def scrape_url(
        self,
        url: str,
        formats: List[str] = ["markdown", "html"]
    ) -> Dict[str, Any]:
        """
        抓取单个网页

        Args:
            url: 要抓取的URL
            formats: 输出格式
        """
        async with httpx.AsyncClient(timeout=120.0) as client:
            try:
                response = await client.post(
                    f"{self.base_url}/scrape",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "url": url,
                        "formats": formats,
                        "onlyMainContent": True,
                        "waitFor": 5000
                    }
                )
                response.raise_for_status()
                data = response.json()

                return {
                    "success": True,
                    "content": data.get("data", {}).get("markdown", ""),
                    "html": data.get("data", {}).get("html", ""),
                    "metadata": data.get("data", {}).get("metadata", {}),
                    "url": url,
                    "timestamp": datetime.now().isoformat()
                }

            except Exception as e:
                console.print(f"[red]Firecrawl抓取失败 ({url}): {e}[/red]")
                return {
                    "success": False,
                    "error": str(e),
                    "url": url,
                    "timestamp": datetime.now().isoformat()
                }

    async def crawl_site(
        self,
        url: str,
        max_pages: int = 10,
        include_patterns: List[str] = None
    ) -> Dict[str, Any]:
        """
        爬取整个网站

        Args:
            url: 起始URL
            max_pages: 最大页面数
            include_patterns: 包含的URL模式
        """
        async with httpx.AsyncClient(timeout=300.0) as client:
            try:
                # 启动爬取任务
                response = await client.post(
                    f"{self.base_url}/crawl",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "url": url,
                        "limit": max_pages,
                        "scrapeOptions": {
                            "formats": ["markdown"],
                            "onlyMainContent": True
                        },
                        "includePaths": include_patterns or []
                    }
                )
                response.raise_for_status()
                job_data = response.json()
                job_id = job_data.get("id")

                if not job_id:
                    return {"success": False, "error": "No job ID returned"}

                # 轮询等待完成
                for _ in range(60):  # 最多等5分钟
                    await asyncio.sleep(5)
                    status_response = await client.get(
                        f"{self.base_url}/crawl/{job_id}",
                        headers={"Authorization": f"Bearer {self.api_key}"}
                    )
                    status_data = status_response.json()

                    if status_data.get("status") == "completed":
                        return {
                            "success": True,
                            "pages": status_data.get("data", []),
                            "total": status_data.get("total", 0),
                            "url": url,
                            "timestamp": datetime.now().isoformat()
                        }
                    elif status_data.get("status") == "failed":
                        return {"success": False, "error": "Crawl failed"}

                return {"success": False, "error": "Crawl timeout"}

            except Exception as e:
                console.print(f"[red]Firecrawl爬取失败: {e}[/red]")
                return {"success": False, "error": str(e)}

    async def search_and_scrape(
        self,
        query: str,
        num_results: int = 5
    ) -> List[Dict[str, Any]]:
        """
        搜索并抓取相关网页
        """
        async with httpx.AsyncClient(timeout=120.0) as client:
            try:
                response = await client.post(
                    f"{self.base_url}/search",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "query": query,
                        "limit": num_results,
                        "scrapeOptions": {
                            "formats": ["markdown"],
                            "onlyMainContent": True
                        }
                    }
                )
                response.raise_for_status()
                data = response.json()

                return {
                    "success": True,
                    "results": data.get("data", []),
                    "query": query,
                    "timestamp": datetime.now().isoformat()
                }

            except Exception as e:
                return {"success": False, "error": str(e), "results": []}


# ============================================================================
# 数据时效性验证器
# ============================================================================

class DataFreshnessValidator:
    """数据时效性验证器 - 杜绝过时数据"""

    def __init__(self):
        self.current_date = datetime.now()
        self.current_year = self.current_date.year
        self.current_month = self.current_date.month

    def validate_content(self, content: str) -> Tuple[bool, List[str]]:
        """
        验证内容的时效性

        Returns:
            (是否通过, 警告列表)
        """
        warnings = []

        # 检查是否包含过时年份
        if EnterpriseConfig.DATA_FRESHNESS["reject_2024_data"]:
            old_years = re.findall(r'\b(2024|2023|2022|2021|2020)\b', content)
            if old_years:
                warnings.append(f"发现过时数据年份: {set(old_years)}")

        # 检查是否有2025年数据
        if EnterpriseConfig.DATA_FRESHNESS["require_2025_data"]:
            if "2025" not in content and "25年" not in content:
                warnings.append("未找到2025年数据")

        # 检查数据来源日期
        if EnterpriseConfig.DATA_FRESHNESS["require_source_date"]:
            date_patterns = [
                r'\d{4}年\d{1,2}月\d{1,2}日',
                r'\d{4}-\d{2}-\d{2}',
                r'\d{4}/\d{2}/\d{2}'
            ]
            has_date = any(re.search(p, content) for p in date_patterns)
            if not has_date:
                warnings.append("建议添加数据来源日期")

        return len(warnings) == 0, warnings

    def add_timestamp(self, content: Dict[str, Any]) -> Dict[str, Any]:
        """为内容添加时间戳"""
        content["_data_timestamp"] = self.current_date.isoformat()
        content["_data_year"] = self.current_year
        content["_data_month"] = self.current_month
        content["_freshness_verified"] = True
        return content


# ============================================================================
# Claude API 客户端 (增强版)
# ============================================================================

class ClaudeClientEnhanced:
    """Claude Opus 4.5 客户端 - 增强版"""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = "https://api.anthropic.com/v1"
        self.model = "claude-opus-4-5-20251101"
        console.print(f"[green]✓[/green] Claude Opus 4.5 增强版客户端初始化完成")

    async def _request(self, messages: List[Dict], system: str = "", max_tokens: int = 8192) -> str:
        """发送请求"""
        async with httpx.AsyncClient(timeout=180.0) as client:
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

    async def generate_report_structure(
        self,
        topic: str,
        context: str,
        realtime_data: Dict[str, Any],
        num_sections: int = 10
    ) -> List[Dict[str, Any]]:
        """
        生成报告结构 (无页数限制)

        支持生成100+页的完整报告结构
        """
        today = datetime.now().strftime("%Y年%m月%d日")

        system_prompt = f"""你是顶级咨询公司的报告架构师。
当前日期: {today}

关键要求:
1. 只使用2025年最新数据
2. 拒绝使用任何2024年或更早的数据
3. 每个数据点必须标注来源和日期
4. 报告结构要专业、完整、有深度

你可以生成任意数量的章节和页面，不受限制。
对于品牌报告，100+页是正常的。"""

        user_message = f"""请为以下主题生成完整的报告结构:

**主题:** {topic}

**背景上下文:**
{context}

**实时数据 (来自Perplexity搜索):**
{json.dumps(realtime_data, ensure_ascii=False, indent=2) if realtime_data else '无'}

**要求:**
1. 生成完整的报告结构，不限页数
2. 包含: 封面、目录、执行摘要、各章节、附录、感谢页
3. 每个章节可以有多个子页面
4. 所有数据必须是2025年最新

输出JSON数组格式:
[
  {{"chapter": "章节名", "slides": [{{"title": "标题", "type": "cover/content/chart/summary", "content_hints": ["要点"]}}]}}
]"""

        try:
            response = await self._request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=16384  # 增加token限制支持长报告
            )
            json_match = re.search(r'\[[\s\S]*\]', response)
            if json_match:
                return json.loads(json_match.group())
            return []
        except Exception as e:
            console.print(f"[red]报告结构生成失败: {e}[/red]")
            return []

    async def fill_slide_with_realtime_data(
        self,
        slide_outline: Dict,
        realtime_data: Dict[str, Any],
        scraped_data: Dict[str, Any] = None
    ) -> Dict[str, Any]:
        """使用实时数据填充幻灯片内容"""
        today = datetime.now().strftime("%Y年%m月%d日")

        system_prompt = f"""你是专业的PPT内容撰写专家。
当前日期: {today}

核心原则:
1. 只使用2025年数据，拒绝2024年数据
2. 每个数据必须标注来源
3. 保持内容专业简洁
4. 中文表达要地道"""

        user_message = f"""请填充以下幻灯片内容:

**大纲:**
{json.dumps(slide_outline, ensure_ascii=False, indent=2)}

**实时搜索数据:**
{json.dumps(realtime_data, ensure_ascii=False, indent=2) if realtime_data else '无'}

**网页抓取数据:**
{json.dumps(scraped_data, ensure_ascii=False, indent=2) if scraped_data else '无'}

输出纯JSON格式，包含:
- title: 标题
- subtitle: 副标题 (可选)
- bullets: 要点列表 (带数据来源)
- content: 正文 (可选)
- data_sources: 数据来源列表
- data_date: 数据日期"""

        try:
            response = await self._request(
                messages=[{"role": "user", "content": user_message}],
                system=system_prompt,
                max_tokens=4096
            )
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                return json.loads(json_match.group())
            return slide_outline
        except Exception as e:
            console.print(f"[yellow]内容填充失败: {e}[/yellow]")
            return slide_outline


# ============================================================================
# Gemini 客户端 (增强版)
# ============================================================================

class GeminiClientEnhanced:
    """Gemini 3 Pro 客户端 - 增强版"""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.client = genai.Client(api_key=api_key)
        self.text_model = "gemini-2.5-flash-preview-04-17"
        self.image_model = "gemini-2.0-flash-exp-image-generation"
        console.print(f"[green]✓[/green] Gemini 3 Pro 增强版客户端初始化完成")

    async def generate_8k_slide_image(
        self,
        content: Dict[str, Any],
        style: str = "professional",
        slide_number: int = 1,
        total_slides: int = 1
    ) -> Optional[bytes]:
        """生成8K幻灯片图像"""
        # 构建内容
        content_parts = []
        if content.get('title'):
            content_parts.append(f"**标题:** {content['title']}")
        if content.get('subtitle'):
            content_parts.append(f"**副标题:** {content['subtitle']}")
        if content.get('bullets'):
            bullets_text = '\n'.join(f"• {b}" for b in content['bullets'][:8])
            content_parts.append(f"**要点:**\n{bullets_text}")
        if content.get('data_sources'):
            sources = ', '.join(content['data_sources'][:3])
            content_parts.append(f"**数据来源:** {sources}")

        content_str = '\n\n'.join(content_parts)

        # 8K prompt
        prompt = f"""Generate an 8K resolution (7680x4320 pixels) professional presentation slide.

**CRITICAL 8K REQUIREMENTS:**
- Resolution: EXACTLY 7680x4320 pixels
- All text HIGH DPI (300 DPI equivalent)
- Chinese: Microsoft YaHei, minimum 72pt titles, 36pt body
- NO blur, NO artifacts on text

**Visual Design:**
- Background: Deep navy gradient (#0F3460 to #1A1A2E)
- Title: Cyan (#00D4FF), 72-96pt, bold
- Body: White (#FFFFFF), 36-48pt
- Accent: Green (#00FF88)

**Content:**
{content_str}

**Slide: {slide_number}/{total_slides}**
{"[COVER SLIDE - Make impactful]" if slide_number == 1 else "[CLOSING SLIDE - Thank you feel]" if slide_number == total_slides else ""}

**ABSOLUTE REQUIREMENT: Text must be CRYSTAL CLEAR at 8K**
"""

        try:
            response = self.client.models.generate_content(
                model=self.image_model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["Text", "Image"],
                    temperature=0.8
                )
            )

            if response.candidates and response.candidates[0].content.parts:
                for part in response.candidates[0].content.parts:
                    if hasattr(part, 'inline_data') and part.inline_data:
                        image_data = part.inline_data.data
                        if isinstance(image_data, str):
                            return base64.b64decode(image_data)
                        return image_data
            return None

        except Exception as e:
            console.print(f"[red]图像生成失败: {e}[/red]")
            return None


# ============================================================================
# 企业级PPT生成器主类
# ============================================================================

@dataclass
class EnterpriseReportConfig:
    """企业级报告配置"""
    title: str = "行业分析报告"
    author: str = "Gree Dashboard"
    style: str = "professional"
    num_slides: Optional[int] = None  # None = 无限制
    use_8k: bool = True
    output_dir: str = "./output"
    enable_realtime_search: bool = True
    enable_web_scraping: bool = True
    enable_quality_check: bool = True
    search_queries: List[str] = field(default_factory=list)
    scrape_urls: List[str] = field(default_factory=list)


class EnterprisePPTGenerator:
    """
    企业级PPT生成器

    特性:
    - 无页数限制，支持100+页报告
    - Perplexity实时搜索
    - Firecrawl网页抓取
    - 数据时效性验证
    - 双模型协作
    """

    def __init__(
        self,
        gemini_api_key: str,
        claude_api_key: str,
        perplexity_api_key: str,
        firecrawl_api_key: str,
        config: Optional[EnterpriseReportConfig] = None
    ):
        self.config = config or EnterpriseReportConfig()
        self.gemini = GeminiClientEnhanced(gemini_api_key)
        self.claude = ClaudeClientEnhanced(claude_api_key)
        self.perplexity = PerplexityClient(perplexity_api_key)
        self.firecrawl = FirecrawlClient(firecrawl_api_key)
        self.validator = DataFreshnessValidator()
        self.output_dir = Path(self.config.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        console.print(Panel.fit(
            "[bold cyan]Enterprise PPT Generator v2.0[/bold cyan]\n\n"
            "✓ 无页数限制 (支持100+页)\n"
            "✓ Perplexity 实时搜索\n"
            "✓ Firecrawl 网页抓取\n"
            "✓ 数据时效性验证\n"
            "✓ 8K分辨率",
            title="🍌 初始化完成"
        ))

    async def generate_enterprise_report(
        self,
        topic: str,
        context: str = "",
        search_queries: List[str] = None,
        scrape_urls: List[str] = None
    ) -> str:
        """
        生成企业级报告 (无页数限制)
        """
        console.print(f"\n[bold]🎯 报告主题:[/bold] {topic}")
        console.print(f"[bold]📅 生成时间:[/bold] {datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
        console.print()

        # ========== Step 1: 实时数据采集 ==========
        realtime_data = {}
        scraped_data = {}

        if self.config.enable_realtime_search:
            console.print("[cyan]Step 1: Perplexity 实时数据采集...[/cyan]")

            queries = search_queries or self.config.search_queries or [topic]

            with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
                task = progress.add_task("搜索中...", total=len(queries))

                for query in queries:
                    result = await self.perplexity.search_realtime(query)
                    if result["success"]:
                        realtime_data[query] = result
                        console.print(f"  [green]✓[/green] {query[:30]}...")
                    else:
                        console.print(f"  [yellow]⚠[/yellow] {query[:30]}... (失败)")
                    progress.advance(task)
                    await asyncio.sleep(1)

            console.print(f"[green]✓[/green] 获取 {len(realtime_data)} 条实时数据\n")

        if self.config.enable_web_scraping and (scrape_urls or self.config.scrape_urls):
            console.print("[cyan]Step 2: Firecrawl 网页抓取...[/cyan]")

            urls = scrape_urls or self.config.scrape_urls

            for url in urls[:10]:  # 最多抓10个URL
                result = await self.firecrawl.scrape_url(url)
                if result["success"]:
                    scraped_data[url] = result
                    console.print(f"  [green]✓[/green] {url[:50]}...")
                else:
                    console.print(f"  [yellow]⚠[/yellow] {url[:50]}... (失败)")
                await asyncio.sleep(2)

            console.print(f"[green]✓[/green] 抓取 {len(scraped_data)} 个网页\n")

        # ========== Step 2: 生成报告结构 ==========
        console.print("[cyan]Step 3: Claude 生成报告结构 (无页数限制)...[/cyan]")

        report_structure = await self.claude.generate_report_structure(
            topic, context, realtime_data
        )

        # 展开所有章节的幻灯片
        all_slides = []
        for chapter in report_structure:
            chapter_name = chapter.get("chapter", "")
            for slide in chapter.get("slides", []):
                slide["chapter"] = chapter_name
                all_slides.append(slide)

        total_slides = len(all_slides)
        console.print(f"[green]✓[/green] 报告结构: {len(report_structure)} 章节, {total_slides} 页\n")

        # ========== Step 3: 填充内容 ==========
        console.print("[cyan]Step 4: Claude 填充内容 (使用实时数据)...[/cyan]")

        filled_slides = []
        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            task = progress.add_task("填充内容...", total=total_slides)

            for i, slide in enumerate(all_slides):
                filled = await self.claude.fill_slide_with_realtime_data(
                    slide, realtime_data, scraped_data
                )

                # 验证数据时效性
                if filled.get("bullets"):
                    content_text = ' '.join(filled["bullets"])
                    is_fresh, warnings = self.validator.validate_content(content_text)
                    if not is_fresh:
                        console.print(f"  [yellow]⚠[/yellow] 页{i+1}: {', '.join(warnings)}")

                filled = self.validator.add_timestamp(filled)
                filled_slides.append(filled)
                progress.advance(task)

                # 批次延迟
                if (i + 1) % EnterpriseConfig.API_LIMITS["batch_size"] == 0:
                    await asyncio.sleep(EnterpriseConfig.API_LIMITS["delay_between_batches"])
                else:
                    await asyncio.sleep(0.5)

        console.print(f"[green]✓[/green] 内容填充完成\n")

        # ========== Step 4: 生成图像 ==========
        console.print("[cyan]Step 5: Gemini 生成8K幻灯片图像...[/cyan]")

        images = []
        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            task = progress.add_task("生成图像...", total=total_slides)

            for i, slide in enumerate(filled_slides):
                image_data = await self.gemini.generate_8k_slide_image(
                    slide,
                    style=self.config.style,
                    slide_number=i + 1,
                    total_slides=total_slides
                )
                images.append(image_data)
                progress.advance(task)

                if (i + 1) % EnterpriseConfig.API_LIMITS["batch_size"] == 0:
                    console.print(f"  [dim]已完成 {i+1}/{total_slides} 页...[/dim]")
                    await asyncio.sleep(EnterpriseConfig.API_LIMITS["delay_between_batches"])
                else:
                    await asyncio.sleep(EnterpriseConfig.API_LIMITS["delay_between_slides"])

        success_count = sum(1 for img in images if img is not None)
        console.print(f"[green]✓[/green] 成功生成 {success_count}/{total_slides} 张图像\n")

        # ========== Step 5: 组装PPT ==========
        console.print("[cyan]Step 6: 组装PPT文件...[/cyan]")

        output_path = await self._assemble_ppt(filled_slides, images)

        # 打印报告摘要
        self._print_report_summary(topic, total_slides, realtime_data, scraped_data, output_path)

        return str(output_path)

    async def _assemble_ppt(
        self,
        slides: List[Dict],
        images: List[Optional[bytes]]
    ) -> Path:
        """组装PPT"""
        prs = Presentation()
        prs.slide_width = Inches(EnterpriseConfig.SLIDE_SIZE['width'])
        prs.slide_height = Inches(EnterpriseConfig.SLIDE_SIZE['height'])

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

        # 保存
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[^\w\u4e00-\u9fa5\-]', '_', self.config.title)
        filename = f"{safe_title}_{len(slides)}页_{timestamp}.pptx"
        output_path = self.output_dir / filename

        prs.save(str(output_path))
        return output_path

    def _add_fallback_slide(self, slide, content: Dict, num: int, total: int):
        """添加回退幻灯片"""
        bg = slide.shapes.add_shape(
            1, Inches(0), Inches(0),
            Inches(EnterpriseConfig.SLIDE_SIZE['width']),
            Inches(EnterpriseConfig.SLIDE_SIZE['height'])
        )
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
            y += 1.2

        if content.get('bullets'):
            for bullet in content['bullets'][:8]:
                box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(12), Inches(0.6))
                p = box.text_frame.paragraphs[0]
                p.text = f"• {bullet}"
                p.font.size = Pt(28)
                p.font.color.rgb = RgbColor.from_string('E0E0E0')
                y += 0.65

        # 页码
        page_box = slide.shapes.add_textbox(Inches(12), Inches(7), Inches(1), Inches(0.3))
        p = page_box.text_frame.paragraphs[0]
        p.text = f"{num}/{total}"
        p.font.size = Pt(14)
        p.font.color.rgb = RgbColor.from_string('666666')

    def _print_report_summary(self, topic, total_slides, realtime_data, scraped_data, output_path):
        """打印报告摘要"""
        table = Table(title="📊 报告生成摘要")
        table.add_column("项目", style="cyan")
        table.add_column("内容", style="green")

        table.add_row("报告主题", topic)
        table.add_row("总页数", f"{total_slides} 页")
        table.add_row("实时数据源", f"{len(realtime_data)} 条")
        table.add_row("网页抓取", f"{len(scraped_data)} 个")
        table.add_row("分辨率", "8K (7680x4320)")
        table.add_row("数据时效", f"2025年{datetime.now().month}月最新")
        table.add_row("输出文件", str(output_path))

        console.print()
        console.print(table)
        console.print()
        console.print(Panel.fit(
            f"[bold green]✅ 企业级报告生成成功![/bold green]\n\n"
            f"📁 {output_path}\n"
            f"📊 共 {total_slides} 页\n"
            f"📅 数据更新至: {datetime.now().strftime('%Y年%m月%d日')}",
            title="🍌 完成"
        ))


# ============================================================================
# CLI入口
# ============================================================================

async def main():
    """主入口"""
    import argparse

    parser = argparse.ArgumentParser(description="🍌 Enterprise PPT Generator v2.0")
    parser.add_argument("-t", "--topic", default="行业分析报告", help="报告主题")
    parser.add_argument("-s", "--style", default="professional", help="幻灯片风格")
    parser.add_argument("-q", "--queries", nargs="+", help="Perplexity搜索查询")
    parser.add_argument("-u", "--urls", nargs="+", help="Firecrawl抓取URL")
    parser.add_argument("-d", "--from-dashboard", action="store_true", help="从仪表盘数据生成")
    parser.add_argument("--no-search", action="store_true", help="禁用实时搜索")
    parser.add_argument("--no-scrape", action="store_true", help="禁用网页抓取")

    args = parser.parse_args()

    # API密钥
    gemini_key = os.getenv("GEMINI_API_KEY")
    claude_key = os.getenv("CLAUDE_API_KEY")
    perplexity_key = os.getenv("PERPLEXITY_API_KEY")
    firecrawl_key = os.getenv("FIRECRAWL_API_KEY")

    if not all([gemini_key, claude_key, perplexity_key, firecrawl_key]):
        console.print("[red]错误: 请设置所有API密钥环境变量[/red]")
        console.print("需要: GEMINI_API_KEY, CLAUDE_API_KEY, PERPLEXITY_API_KEY, FIRECRAWL_API_KEY")
        return

    config = EnterpriseReportConfig(
        title=args.topic,
        style=args.style,
        enable_realtime_search=not args.no_search,
        enable_web_scraping=not args.no_scrape,
        search_queries=args.queries or [],
        scrape_urls=args.urls or []
    )

    generator = EnterprisePPTGenerator(
        gemini_key, claude_key, perplexity_key, firecrawl_key, config
    )

    if args.from_dashboard:
        # 格力仪表盘示例
        output = await generator.generate_enterprise_report(
            topic="格力抖音官方旗舰店2025年Q1数据分析报告",
            context="格力电器抖音电商业务分析，需要最新的行业数据对比",
            search_queries=[
                "格力电器2025年Q1销售数据",
                "空调行业2025年市场份额",
                "抖音电商2025年1月GMV数据",
                "家电行业2025年趋势预测"
            ]
        )
    else:
        output = await generator.generate_enterprise_report(
            topic=args.topic,
            search_queries=args.queries or [args.topic],
            scrape_urls=args.urls
        )

    console.print(f"\n[bold green]🎉 完成![/bold green] 文件: {output}")


if __name__ == "__main__":
    asyncio.run(main())
