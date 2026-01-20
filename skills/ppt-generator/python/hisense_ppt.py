#!/usr/bin/env python3
"""
海信AGI直播运营系统PPT生成器
直接生成30页专业商业提案PPT
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import os

# 配色方案
COLORS = {
    'bg_dark': '0d1117',      # 深黑背景
    'primary': '00d97e',       # 科技绿
    'blue': '58a6ff',          # 辅助蓝
    'orange': 'f0883e',        # 警告橙
    'red': 'f85149',           # 危险红
    'card_bg': '161b22',       # 卡片背景
    'border': '30363d',        # 边框
    'text_main': 'e6edf3',     # 主文字
    'text_secondary': '8b949e' # 次要文字
}

def hex_to_rgbcolor(hex_color):
    return RGBColor.from_string(hex_color)

def add_background(slide, color='0d1117'):
    """添加深色背景"""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgbcolor(color)
    bg.line.fill.background()
    # 移到最底层
    spTree = slide.shapes._spTree
    sp = bg._element
    spTree.remove(sp)
    spTree.insert(2, sp)

def add_left_accent(slide, color='00d97e', width=0.08):
    """添加左侧强调条"""
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        Inches(width), Inches(7.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = hex_to_rgbcolor(color)
    accent.line.fill.background()

def add_title(slide, title, subtitle=None, y=0.5):
    """添加标题"""
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(12), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('e6edf3')

    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(y + 0.7), Inches(12), Inches(0.5))
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        p.text = subtitle
        p.font.size = Pt(18)
        p.font.color.rgb = hex_to_rgbcolor('8b949e')

def add_section_tag(slide, text, color='00d97e'):
    """添加章节标签"""
    tag = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(0.3),
        Inches(3), Inches(0.35)
    )
    tag.fill.solid()
    tag.fill.fore_color.rgb = hex_to_rgbcolor(color)
    tag.fill.fore_color.brightness = 0.8
    tag.line.fill.background()

    tag_text = slide.shapes.add_textbox(Inches(0.5), Inches(0.32), Inches(3), Inches(0.35))
    tf = tag_text.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor(color)
    p.alignment = PP_ALIGN.CENTER

def add_card(slide, x, y, w, h, title, content, icon='', border_color='30363d', title_color='e6edf3'):
    """添加卡片"""
    # 卡片背景
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y),
        Inches(w), Inches(h)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    card.line.color.rgb = hex_to_rgbcolor(border_color)
    card.line.width = Pt(1.5)

    # 标题
    title_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.15), Inches(w - 0.4), Inches(0.4))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"{icon} {title}" if icon else title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor(title_color)

    # 内容
    content_box = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.55), Inches(w - 0.4), Inches(h - 0.7))
    tf = content_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = content
    p.font.size = Pt(12)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

def add_big_number(slide, x, y, number, label, color='00d97e'):
    """添加大数字"""
    num_box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(3), Inches(0.8))
    tf = num_box.text_frame
    p = tf.paragraphs[0]
    p.text = number
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor(color)

    label_box = slide.shapes.add_textbox(Inches(x), Inches(y + 0.7), Inches(3), Inches(0.4))
    tf = label_box.text_frame
    p = tf.paragraphs[0]
    p.text = label
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

def add_progress_bar(slide, x, y, w, value, max_val, label, color='00d97e'):
    """添加进度条"""
    # 背景条
    bg_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y + 0.3),
        Inches(w), Inches(0.15)
    )
    bg_bar.fill.solid()
    bg_bar.fill.fore_color.rgb = hex_to_rgbcolor('30363d')
    bg_bar.line.fill.background()

    # 进度条
    progress_w = w * (value / max_val)
    if progress_w > 0.1:
        prog_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y + 0.3),
            Inches(progress_w), Inches(0.15)
        )
        prog_bar.fill.solid()
        prog_bar.fill.fore_color.rgb = hex_to_rgbcolor(color)
        prog_bar.line.fill.background()

    # 标签
    label_box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w * 0.6), Inches(0.3))
    tf = label_box.text_frame
    p = tf.paragraphs[0]
    p.text = label
    p.font.size = Pt(11)
    p.font.color.rgb = hex_to_rgbcolor('e6edf3')

    # 数值
    val_box = slide.shapes.add_textbox(Inches(x + w * 0.7), Inches(y), Inches(w * 0.3), Inches(0.3))
    tf = val_box.text_frame
    p = tf.paragraphs[0]
    p.text = str(value)
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor(color)
    p.alignment = PP_ALIGN.RIGHT

def create_hisense_ppt():
    """创建海信AGI直播运营系统PPT"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # ==================== P1 封面 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)

    # 顶部渐变条
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.1))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = hex_to_rgbcolor('00d97e')
    top_bar.line.fill.background()

    # AGI标签
    tag = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(3), Inches(0.4))
    tf = tag.text_frame
    p = tf.paragraphs[0]
    p.text = "AGI LIVEOPS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # 主标题
    title = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12), Inches(1))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "直播管理的「工业革命」"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('e6edf3')

    # 副标题
    subtitle = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(12), Inches(0.6))
    tf = subtitle.text_frame
    p = tf.paragraphs[0]
    p.text = "从人力堆叠到 AGI 算力驱动"
    p.font.size = Pt(24)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # 分隔线
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4.3), Inches(1.5), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgbcolor('00d97e')
    line.line.fill.background()

    # 理念文案
    slogan = slide.shapes.add_textbox(Inches(0.5), Inches(4.6), Inches(12), Inches(0.5))
    tf = slogan.text_frame
    p = tf.paragraphs[0]
    p.text = "让运营回归本质，让增长有迹可循"
    p.font.size = Pt(18)
    p.font.color.rgb = hex_to_rgbcolor('58a6ff')

    # 底部署名
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(12), Inches(0.4))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.text = "方案独家提供方：聚能鼎力 AGI 实验室"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P2 现状拷问 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_section_tag(slide, "PART 01 · THE INSIGHT")
    add_title(slide, "数字化深水区：为何千万投入，依然难解「实时」之痛？", y=0.8)

    # 三张卡片
    add_card(slide, 0.5, 1.8, 4, 2.5, "业务门槛极高",
             "直播不是标准化流水线，是高度非标的「情绪博弈场」。\n\n纯技术思维无法理解业务内核。\n\n海信花了1000万+找淘宝团队，依然没搞定。",
             "⚠️", 'f85149', 'f85149')

    add_card(slide, 4.7, 1.8, 4, 2.5, "单一驱动失效",
             "光有业务经验 = 无法规模化复制\n\n光有技术数据 = T+1 的滞后报表\n\n（后视镜开车，看到的都是过去）",
             "⚡", 'f0883e', 'f0883e')

    add_card(slide, 8.9, 1.8, 4, 2.5, "双轮驱动",
             "未来核心壁垒：\n\n顶级操盘手 Know-How\n×\n多模态 AGI 技术\n\n这是我们的独家护城河。",
             "🎯", '00d97e', '00d97e')

    # 底部对比
    bottom = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(12), Inches(1.5))
    tf = bottom.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "行业现状：90%的直播运营系统只能做T+1日报，等你看到数据，流量早跑了。\n我们的方案：T+0 毫秒级响应，流量还在的时候就把问题解决。"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P3 重新定义业务 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, "直播带货 = 瞬间的「情绪变现」", y=0.5)

    # 大数字卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(5), Inches(3))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    card.line.color.rgb = hex_to_rgbcolor('00d97e')
    card.line.width = Pt(2)

    add_big_number(slide, 1.5, 2, "1秒", "决策时间窗口", '00d97e')

    desc = slide.shapes.add_textbox(Inches(0.8), Inches(3.5), Inches(4.4), Inches(0.8))
    tf = desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "消费者下单往往就在那一秒的冲动。这是直播电商的核心法则。"
    p.font.size = Pt(13)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # 右侧卡片
    add_card(slide, 6, 1.5, 6.8, 1.4, "🔍 AGI 价值",
             "捕捉这「关键一秒」：主播眼神是否坚定？声音是否高亢？痛点是否戳中？",
             '', '00d97e', '00d97e')

    add_card(slide, 6, 3.1, 6.8, 1.4, "💡 核心结论",
             "这是 ERP 抓不到的数据，是视觉大模型抓到的「人性」。",
             '', '58a6ff', '58a6ff')

    # 底部
    bottom = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(12), Inches(1.5))
    tf = bottom.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "传统方案的致命缺陷：只看销售数据，不看销售过程。等GMV出来再复盘，黄花菜都凉了。"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('f0883e')

    # ==================== P4 技术降维 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, "AGI 的「上帝视角」", y=0.5)

    # 人工场控卡片
    card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(6), Inches(3.5))
    card1.fill.solid()
    card1.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    card1.line.color.rgb = hex_to_rgbcolor('f85149')
    card1.line.width = Pt(2)

    t1 = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(5.4), Inches(0.5))
    tf = t1.text_frame
    p = tf.paragraphs[0]
    p.text = "❌ 人工场控"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('f85149')

    c1 = slide.shapes.add_textbox(Inches(0.8), Inches(2.3), Inches(5.4), Inches(2.5))
    tf = c1.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• 只能听懂「说了什么」（内容）\n\n• 会走神、带情绪\n\n• 一个人最多盯2个直播间\n\n• 下班了就没人盯了\n\n• 主观判断，标准不一"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # AGI场控卡片
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(6), Inches(3.5))
    card2.fill.solid()
    card2.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    card2.line.color.rgb = hex_to_rgbcolor('00d97e')
    card2.line.width = Pt(2)

    t2 = slide.shapes.add_textbox(Inches(7.1), Inches(1.7), Inches(5.4), Inches(0.5))
    tf = t2.text_frame
    p = tf.paragraphs[0]
    p.text = "✅ AGI 场控"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    c2 = slide.shapes.add_textbox(Inches(7.1), Inches(2.3), Inches(5.4), Inches(2.5))
    tf = c2.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "• 能听懂「怎么说的」（情绪、语气、潜台词）\n\n• 24小时毫秒级精度在线\n\n• 同时监控100+直播间\n\n• 永不疲劳，永不下班\n\n• 统一标准，量化考核"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # 底部结论
    bottom = slide.shapes.add_textbox(Inches(0.5), Inches(5.3), Inches(12), Inches(0.8))
    tf = bottom.text_frame
    p = tf.paragraphs[0]
    p.text = '降维打击：从"内容审核"到"情绪感知"'
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # ==================== P5 系统定位 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, "不是工具，是「超级大脑」", y=0.5)

    # 左侧圆形 - 工具
    circle1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(2.5), Inches(2.5))
    circle1.fill.solid()
    circle1.fill.fore_color.rgb = hex_to_rgbcolor('30363d')
    circle1.line.color.rgb = hex_to_rgbcolor('f85149')
    circle1.line.width = Pt(3)

    t1 = slide.shapes.add_textbox(Inches(2), Inches(2.8), Inches(2.5), Inches(1))
    tf = t1.text_frame
    p = tf.paragraphs[0]
    p.text = "📊\n工具"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('f85149')
    p.alignment = PP_ALIGN.CENTER

    label1 = slide.shapes.add_textbox(Inches(2), Inches(4.7), Inches(2.5), Inches(0.4))
    tf = label1.text_frame
    p = tf.paragraphs[0]
    p.text = "记录数据 · 被动响应"
    p.font.size = Pt(12)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')
    p.alignment = PP_ALIGN.CENTER

    # 箭头
    arrow = slide.shapes.add_textbox(Inches(5), Inches(3), Inches(1), Inches(0.8))
    tf = arrow.text_frame
    p = tf.paragraphs[0]
    p.text = "→"
    p.font.size = Pt(48)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')
    p.alignment = PP_ALIGN.CENTER

    # 右侧圆形 - 大脑
    circle2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.5), Inches(2), Inches(2.5), Inches(2.5))
    circle2.fill.solid()
    circle2.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    circle2.line.color.rgb = hex_to_rgbcolor('00d97e')
    circle2.line.width = Pt(3)

    t2 = slide.shapes.add_textbox(Inches(8.5), Inches(2.8), Inches(2.5), Inches(1))
    tf = t2.text_frame
    p = tf.paragraphs[0]
    p.text = "🧠\n大脑"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')
    p.alignment = PP_ALIGN.CENTER

    label2 = slide.shapes.add_textbox(Inches(8.5), Inches(4.7), Inches(2.5), Inches(0.4))
    tf = label2.text_frame
    p = tf.paragraphs[0]
    p.text = "制造决策 · 主动出击"
    p.font.size = Pt(12)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')
    p.alignment = PP_ALIGN.CENTER

    # 底部文案
    bottom = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(12), Inches(1))
    tf = bottom.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = '它不是用来「记录」数据的，它是用来「制造」优秀运营决策的。'
    p.font.size = Pt(18)
    p.font.color.rgb = hex_to_rgbcolor('58a6ff')

    # ==================== P6 核心理念 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, "让不确定性变为确定性", y=0.5)

    # 金句卡片
    quote_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1.8), Inches(11.3), Inches(2.5))
    quote_card.fill.solid()
    quote_card.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    quote_card.line.color.rgb = hex_to_rgbcolor('00d97e')
    quote_card.line.width = Pt(2)

    quote = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.3), Inches(1.5))
    tf = quote.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = '"消除人为情绪波动带来的业绩损耗，\n让每一场直播都维持在海信的"金牌标准"。"'
    p.font.size = Pt(24)
    p.font.color.rgb = hex_to_rgbcolor('e6edf3')
    p.alignment = PP_ALIGN.CENTER

    # 三个关键词
    keywords = ["情绪消除", "金牌标准", "确定增长"]
    colors = ['00d97e', '58a6ff', 'f0883e']
    for i, (kw, color) in enumerate(zip(keywords, colors)):
        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3 + i * 2.5), Inches(4.8), Inches(2), Inches(0.5))
        tag.fill.solid()
        tag.fill.fore_color.rgb = hex_to_rgbcolor(color)
        tag.fill.fore_color.brightness = 0.7
        tag.line.fill.background()

        tag_text = slide.shapes.add_textbox(Inches(3 + i * 2.5), Inches(4.85), Inches(2), Inches(0.4))
        tf = tag_text.text_frame
        p = tf.paragraphs[0]
        p.text = kw
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor(color)
        p.alignment = PP_ALIGN.CENTER

    # ==================== P7 双引擎驱动 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_section_tag(slide, "PART 02 · THE CAPABILITIES", '58a6ff')
    add_title(slide, "双引擎驱动 (Vision + Audio)", y=0.8)

    # 三大模块
    modules = [
        ("👁️", "视觉模型", "Eye", "看画面、看表情、看动作", "58a6ff"),
        ("👂", "听觉模型", "Ear", "听声音、听语气、听逻辑", "00d97e"),
        ("✋", "RPA执行", "Hand", "自动操作、实时反馈", "f0883e")
    ]

    for i, (icon, title, eng, desc, color) in enumerate(modules):
        x = 0.8 + i * 4.2

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2), Inches(3.8), Inches(3))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
        card.line.color.rgb = hex_to_rgbcolor(color)
        card.line.width = Pt(2)

        # 图标
        icon_box = slide.shapes.add_textbox(Inches(x), Inches(2.2), Inches(3.8), Inches(0.8))
        tf = icon_box.text_frame
        p = tf.paragraphs[0]
        p.text = icon
        p.font.size = Pt(36)
        p.alignment = PP_ALIGN.CENTER

        # 标题
        title_box = slide.shapes.add_textbox(Inches(x), Inches(3), Inches(3.8), Inches(0.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor(color)
        p.alignment = PP_ALIGN.CENTER

        # 英文
        eng_box = slide.shapes.add_textbox(Inches(x), Inches(3.5), Inches(3.8), Inches(0.4))
        tf = eng_box.text_frame
        p = tf.paragraphs[0]
        p.text = eng
        p.font.size = Pt(12)
        p.font.color.rgb = hex_to_rgbcolor('8b949e')
        p.alignment = PP_ALIGN.CENTER

        # 描述
        desc_box = slide.shapes.add_textbox(Inches(x), Inches(4), Inches(3.8), Inches(0.8))
        tf = desc_box.text_frame
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(14)
        p.font.color.rgb = hex_to_rgbcolor('e6edf3')
        p.alignment = PP_ALIGN.CENTER

    # 底部说明
    bottom = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(12), Inches(0.5))
    tf = bottom.text_frame
    p = tf.paragraphs[0]
    p.text = "三位一体的智能场控系统"
    p.font.size = Pt(16)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')
    p.alignment = PP_ALIGN.CENTER

    # ==================== P8 AI视觉洞察 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, '看见「微表情」里的生意', y=0.5)

    features = [
        ("😴", "状态监测", "识别主播疲劳、眼神游离、假笑（掉粉前兆）"),
        ("✅", "动作规范", '识别是否正确展示「真空保鲜层」、是否遮挡关键卖点'),
        ("💡", "环境诊断", "灯光过暗、背景杂乱、穿帮预警")
    ]

    for i, (icon, title, desc) in enumerate(features):
        add_card(slide, 0.5 + i * 4.2, 1.5, 4, 2.2, f"{icon} {title}", desc, '', '58a6ff', '58a6ff')

    # 行业对比
    compare = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12), Inches(2))
    tf = compare.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "行业现状对比：\n• 传统方案：只能看到销售数字，不知道为什么卖得好/不好\n• 竞品方案：T+1日报，等你发现主播状态差，已经损失了一整场流量\n• 我们的方案：实时检测主播微表情，提前预警，问题发生前就解决"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P9 AI听觉理解 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, '听懂「话术」里的逻辑', y=0.5)

    features = [
        ("📝", "逻辑拆解", '自动判断是否完成了\n「痛点-引入-证言-逼单」的闭环'),
        ("📊", "激情值", '实时分析语调能量\n低于阈值立刻报警（避免冷场）'),
        ("🚫", "违规拦截", '毫秒级识别违禁词\n比平台审核更快一步')
    ]

    for i, (icon, title, desc) in enumerate(features):
        add_card(slide, 0.5 + i * 4.2, 1.5, 4, 2.2, f"{icon} {title}", desc, '', '58a6ff', '58a6ff')

    # 评分示例
    add_progress_bar(slide, 0.5, 4.2, 5, 92, 100, "话术评分", '00d97e')
    add_progress_bar(slide, 6, 4.2, 5, 78, 100, "激情指数", '58a6ff')

    # ==================== P10 T+0实时诊断 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, "T+0 的战术价值", y=0.5)

    # 流程展示
    steps = ["发现问题", "弹窗提醒", "话术修正"]
    labels = ["AI监测", "实时推送", "主播调整"]

    for i, (step, label) in enumerate(zip(steps, labels)):
        x = 1 + i * 4

        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.8), Inches(3), Inches(1.2))
        box.fill.solid()
        box.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
        box.line.color.rgb = hex_to_rgbcolor('00d97e')
        box.line.width = Pt(2)

        step_text = slide.shapes.add_textbox(Inches(x), Inches(2), Inches(3), Inches(0.6))
        tf = step_text.text_frame
        p = tf.paragraphs[0]
        p.text = step
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor('e6edf3')
        p.alignment = PP_ALIGN.CENTER

        label_text = slide.shapes.add_textbox(Inches(x), Inches(2.6), Inches(3), Inches(0.4))
        tf = label_text.text_frame
        p = tf.paragraphs[0]
        p.text = f"[{label}]"
        p.font.size = Pt(12)
        p.font.color.rgb = hex_to_rgbcolor('8b949e')
        p.alignment = PP_ALIGN.CENTER

        if i < 2:
            arrow = slide.shapes.add_textbox(Inches(x + 3), Inches(2.2), Inches(1), Inches(0.5))
            tf = arrow.text_frame
            p = tf.paragraphs[0]
            p.text = "→"
            p.font.size = Pt(24)
            p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # 大数字
    add_big_number(slide, 5, 3.5, "<10秒", "全过程响应时间", '00d97e')

    # 价值陈述
    value = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(12), Inches(0.8))
    tf = value.text_frame
    p = tf.paragraphs[0]
    p.text = "流量还在，单子还能救回来。"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # ==================== P11 赛马机制 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, "数字化赛马机制 - 八大考核维度", y=0.5)

    # 八维评分
    metrics = [
        ("素材数量", 85, '00d97e'),
        ("素材质量", 78, '58a6ff'),
        ("主播话术", 92, '00d97e'),
        ("中控配合", 88, '58a6ff'),
        ("流量爆款", 72, 'f0883e'),
        ("投流ROI", 95, 'f0883e'),
        ("客服转化", 90, '00d97e'),
        ("数据闭环", 82, '58a6ff')
    ]

    for i, (name, score, color) in enumerate(metrics):
        row = i // 4
        col = i % 4
        x = 0.5 + col * 3.2
        y = 1.5 + row * 1.8
        add_progress_bar(slide, x, y, 2.8, score, 100, name, color)

    # 说明
    note = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(12), Inches(1.5))
    tf = note.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "用数据考核，而不是凭关系考核。\n话术分92的上黄金档，话术分60的去淘汰。\n\n这套机制已在截图系统中落地实现（见系统截图-赛马机制八大考核维度）"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P12 竞品监控 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, "实时竞品监控 - 知己知彼", y=0.5)

    # 竞品数据
    competitors = [
        ("海信(本店)", 6, 5251, 1706, '00d97e'),
        ("美的冰箱", 128, 4280, 28600, 'f0883e'),
        ("容声冰箱", 86, 5120, 19800, 'f0883e'),
        ("格力冰箱", None, 6800, 45200, '8b949e'),
        ("海尔冰箱", 215, 7200, 52800, '8b949e')
    ]

    # 表头
    headers = ["店铺", "在线人数", "GPM", "销售额"]
    for i, h in enumerate(headers):
        hbox = slide.shapes.add_textbox(Inches(0.5 + i * 3), Inches(1.5), Inches(2.8), Inches(0.4))
        tf = hbox.text_frame
        p = tf.paragraphs[0]
        p.text = h
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # 数据行
    for j, (name, online, gpm, sales, color) in enumerate(competitors):
        y = 2 + j * 0.7

        # 名称
        n = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(2.8), Inches(0.5))
        tf = n.text_frame
        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(14)
        p.font.color.rgb = hex_to_rgbcolor(color)

        # 在线人数
        o = slide.shapes.add_textbox(Inches(3.5), Inches(y), Inches(2.8), Inches(0.5))
        tf = o.text_frame
        p = tf.paragraphs[0]
        p.text = str(online) if online else "-"
        p.font.size = Pt(14)
        p.font.color.rgb = hex_to_rgbcolor('e6edf3')

        # GPM
        g = slide.shapes.add_textbox(Inches(6.5), Inches(y), Inches(2.8), Inches(0.5))
        tf = g.text_frame
        p = tf.paragraphs[0]
        p.text = f"¥{gpm:,}"
        p.font.size = Pt(14)
        p.font.color.rgb = hex_to_rgbcolor('00d97e' if name == "海信(本店)" else 'e6edf3')

        # 销售额
        s = slide.shapes.add_textbox(Inches(9.5), Inches(y), Inches(2.8), Inches(0.5))
        tf = s.text_frame
        p = tf.paragraphs[0]
        p.text = f"¥{sales:,}"
        p.font.size = Pt(14)
        p.font.color.rgb = hex_to_rgbcolor('e6edf3')

    # 说明
    note = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(12), Inches(1))
    tf = note.text_frame
    p = tf.paragraphs[0]
    p.text = "数据来源：系统截图 - 竞品直播间动态监控（实时抓取抖音罗盘数据）"
    p.font.size = Pt(12)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P13 系统界面展示 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_title(slide, "系统实际界面 - 已落地运行", y=0.5)

    # 四个功能模块
    modules = [
        ("实时监控", "双窗口盯盘：抖音罗盘 + 直播间\n实时在线、支付金额、GPM、新增粉丝"),
        ("违规检测", "主播在镜 ✓ | 仪态正常 ✓\n话术合规 ✓ | 本场违规: 0"),
        ("竞品看板", "海信 vs 美的 vs 容声 vs 格力 vs 海尔\n在线人数、GPM、销售额实时对比"),
        ("赛马机制", "八大维度量化考核\n素材、话术、ROI、转化全覆盖")
    ]

    for i, (title, desc) in enumerate(modules):
        row = i // 2
        col = i % 2
        x = 0.5 + col * 6.4
        y = 1.5 + row * 2.5
        add_card(slide, x, y, 6, 2.2, title, desc, '', '00d97e', '00d97e')

    # 底部说明
    note = slide.shapes.add_textbox(Inches(0.5), Inches(6.2), Inches(12), Inches(0.8))
    tf = note.text_frame
    p = tf.paragraphs[0]
    p.text = "以上功能均已在系统中落地实现，截图可证。不是PPT画饼，是真实可用的系统。"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # ==================== P14 管理水准提升 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_section_tag(slide, "PART 03 · THE VALUE · 管理水准提升")
    add_title(slide, '把「非标」变成「标准」', y=0.8)

    # Before/After对比
    before = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2), Inches(5.5), Inches(2.5))
    before.fill.solid()
    before.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    before.line.color.rgb = hex_to_rgbcolor('f85149')
    before.line.width = Pt(2)

    b_title = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(5), Inches(0.5))
    tf = b_title.text_frame
    p = tf.paragraphs[0]
    p.text = "Before"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('f85149')

    b_num = slide.shapes.add_textbox(Inches(0.8), Inches(2.8), Inches(5), Inches(0.8))
    tf = b_num.text_frame
    p = tf.paragraphs[0]
    p.text = "SOP执行率: 60%"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('f85149')

    b_desc = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), Inches(5), Inches(0.5))
    tf = b_desc.text_frame
    p = tf.paragraphs[0]
    p.text = "SOP是挂在墙上的纸"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # After
    after = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7), Inches(2), Inches(5.5), Inches(2.5))
    after.fill.solid()
    after.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
    after.line.color.rgb = hex_to_rgbcolor('00d97e')
    after.line.width = Pt(2)

    a_title = slide.shapes.add_textbox(Inches(7.3), Inches(2.2), Inches(5), Inches(0.5))
    tf = a_title.text_frame
    p = tf.paragraphs[0]
    p.text = "After"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    a_num = slide.shapes.add_textbox(Inches(7.3), Inches(2.8), Inches(5), Inches(0.8))
    tf = a_num.text_frame
    p = tf.paragraphs[0]
    p.text = "SOP执行率: 100%"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    a_desc = slide.shapes.add_textbox(Inches(7.3), Inches(3.8), Inches(5), Inches(0.5))
    tf = a_desc.text_frame
    p = tf.paragraphs[0]
    p.text = "SOP是AGI实时盯着的红线"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P15 数据颗粒度 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, '从「天」到「秒」', y=0.5)

    # 升级展示
    old = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
    tf = old.text_frame
    p = tf.paragraphs[0]
    p.text = "T+1 日报"
    p.font.size = Pt(28)
    p.font.color.rgb = hex_to_rgbcolor('f85149')

    old_label = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(4), Inches(0.5))
    tf = old_label.text_frame
    p = tf.paragraphs[0]
    p.text = "[后视镜]"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    arrow = slide.shapes.add_textbox(Inches(5), Inches(2.2), Inches(2), Inches(0.8))
    tf = arrow.text_frame
    p = tf.paragraphs[0]
    p.text = "→"
    p.font.size = Pt(48)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    new = slide.shapes.add_textbox(Inches(7), Inches(2), Inches(5), Inches(1))
    tf = new.text_frame
    p = tf.paragraphs[0]
    p.text = "T+0 分钟级心电图"
    p.font.size = Pt(28)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    new_label = slide.shapes.add_textbox(Inches(7), Inches(2.8), Inches(5), Inches(0.5))
    tf = new_label.text_frame
    p = tf.paragraphs[0]
    p.text = "[实时仪表盘]"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    add_big_number(slide, 5, 4, "100倍", "决策精度提升", '00d97e')

    # ==================== P16 人效革命 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, 'f0883e')
    add_section_tag(slide, "PART 03 · THE VALUE · 成本控制", 'f0883e')
    add_title(slide, "1套系统 > 10人团队", y=0.8)

    add_big_number(slide, 5, 2, "60%+", "单直播间人力成本降低", 'f0883e')

    # 可替代岗位
    roles = ["❌ 中控专员", "❌ 数据分析师", "❌ 巡场督导"]
    for i, role in enumerate(roles):
        r = slide.shapes.add_textbox(Inches(1 + i * 4), Inches(4), Inches(3.5), Inches(0.5))
        tf = r.text_frame
        p = tf.paragraphs[0]
        p.text = role
        p.font.size = Pt(18)
        p.font.color.rgb = hex_to_rgbcolor('f85149')

    note = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(12), Inches(1))
    tf = note.text_frame
    p = tf.paragraphs[0]
    p.text = "不是裁员，是让人做更有价值的事。\n海信几百人的直播基地，人效提升意味着每年节省数百万人力成本。"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P17 金牌复制 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '58a6ff')
    add_section_tag(slide, "PART 03 · THE VALUE · 人员定向提升", '58a6ff')
    add_title(slide, '克隆「销冠」的基因', y=0.8)

    # 对比
    add_big_number(slide, 2, 2.2, "30天", "Before: 新人孵化期", 'f85149')

    arrow = slide.shapes.add_textbox(Inches(5.5), Inches(2.5), Inches(2), Inches(0.8))
    tf = arrow.text_frame
    p = tf.paragraphs[0]
    p.text = "→"
    p.font.size = Pt(48)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    add_big_number(slide, 8, 2.2, "3天", "After: AGI辅助", '00d97e')

    desc = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(12), Inches(1.5))
    tf = desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "提取金牌主播的话术结构和节奏，生成模板，让新人照着练。\n\n降幅：90% ↓"
    p.font.size = Pt(16)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P18 能力沉淀 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, 'f85149')
    add_title(slide, "铁打的系统，流水的兵", y=0.5)

    # 对比
    left = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
    tf = left.text_frame
    p = tf.paragraphs[0]
    p.text = "个人能力"
    p.font.size = Pt(24)
    p.font.color.rgb = hex_to_rgbcolor('f85149')

    left_label = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(4), Inches(0.5))
    tf = left_label.text_frame
    p = tf.paragraphs[0]
    p.text = "[会流失]"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    arrow = slide.shapes.add_textbox(Inches(5), Inches(2.2), Inches(2), Inches(0.8))
    tf = arrow.text_frame
    p = tf.paragraphs[0]
    p.text = "→"
    p.font.size = Pt(48)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    right = slide.shapes.add_textbox(Inches(7), Inches(2), Inches(5), Inches(1))
    tf = right.text_frame
    p = tf.paragraphs[0]
    p.text = "组织能力"
    p.font.size = Pt(24)
    p.font.color.rgb = hex_to_rgbcolor('00d97e')

    right_label = slide.shapes.add_textbox(Inches(7), Inches(2.8), Inches(5), Inches(0.5))
    tf = right_label.text_frame
    p = tf.paragraphs[0]
    p.text = "[永久沉淀]"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # 沉淀内容
    items = ["✅ SOP库：永久保存", "✅ 话术库：持续迭代", "✅ 案例库：经验累积"]
    for i, item in enumerate(items):
        t = slide.shapes.add_textbox(Inches(1), Inches(4 + i * 0.5), Inches(11), Inches(0.5))
        tf = t.text_frame
        p = tf.paragraphs[0]
        p.text = item
        p.font.size = Pt(16)
        p.font.color.rgb = hex_to_rgbcolor('00d97e')

    note = slide.shapes.add_textbox(Inches(0.5), Inches(5.8), Inches(12), Inches(0.8))
    tf = note.text_frame
    p = tf.paragraphs[0]
    p.text = "海信不再被大主播绑架，把能力建在组织上。"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('58a6ff')

    # ==================== P19 落地路线图 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_section_tag(slide, "PART 04 · THE VISION")
    add_title(slide, "三步走战略：从接入到自动驾驶", y=0.8)

    phases = [
        ("Phase 1", "2周", "接入期", "数据打通，AGI学习海信知识库", "58a6ff"),
        ("Phase 2", "1个月", "磨合期", "AGI辅助场控，跑通人机协作流程", "00d97e"),
        ("Phase 3", "持续", "进化期", "全自动无人化值守，矩阵化复制推广", "f0883e")
    ]

    for i, (phase, time, name, desc, color) in enumerate(phases):
        x = 0.5 + i * 4.2

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2), Inches(4), Inches(3))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgbcolor('161b22')
        card.line.color.rgb = hex_to_rgbcolor(color)
        card.line.width = Pt(2)

        p_text = slide.shapes.add_textbox(Inches(x + 0.2), Inches(2.2), Inches(3.6), Inches(0.4))
        tf = p_text.text_frame
        p = tf.paragraphs[0]
        p.text = phase
        p.font.size = Pt(12)
        p.font.color.rgb = hex_to_rgbcolor(color)

        t_text = slide.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(3.6), Inches(0.5))
        tf = t_text.text_frame
        p = tf.paragraphs[0]
        p.text = time
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor(color)

        n_text = slide.shapes.add_textbox(Inches(x + 0.2), Inches(3.2), Inches(3.6), Inches(0.4))
        tf = n_text.text_frame
        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgbcolor('e6edf3')

        d_text = slide.shapes.add_textbox(Inches(x + 0.2), Inches(3.7), Inches(3.6), Inches(1))
        tf = d_text.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(12)
        p.font.color.rgb = hex_to_rgbcolor('8b949e')

    # ==================== P20 独家优势 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)
    add_left_accent(slide, '00d97e')
    add_title(slide, "聚能鼎力 AGI 实验室：源头技术，独家赋能", y=0.5)

    # 核心承诺
    promises = [
        "✅ 我们提供的不只是代码，是持续迭代的AGI业务咨询",
        "✅ 本方案核心技术与算力模型，由聚能鼎力AGI实验室独家提供"
    ]

    for i, promise in enumerate(promises):
        p_text = slide.shapes.add_textbox(Inches(0.5), Inches(1.5 + i * 0.6), Inches(12), Inches(0.5))
        tf = p_text.text_frame
        p = tf.paragraphs[0]
        p.text = promise
        p.font.size = Pt(16)
        p.font.color.rgb = hex_to_rgbcolor('00d97e')

    # 独家优势表格
    advantages = [
        ("技术源头", "自研多模态大模型"),
        ("业务理解", "40个月抖音电商实战经验"),
        ("持续迭代", "周级版本更新"),
        ("独家授权", "海信专属定制")
    ]

    for i, (dim, val) in enumerate(advantages):
        row = i // 2
        col = i % 2
        x = 0.5 + col * 6.4
        y = 3 + row * 1.5
        add_card(slide, x, y, 6, 1.2, dim, val, '', '00d97e', '00d97e')

    # ==================== P21 封底 ====================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)

    # 顶部渐变条
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.1))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = hex_to_rgbcolor('00d97e')
    top_bar.line.fill.background()

    # 主文案
    main = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(1))
    tf = main.text_frame
    p = tf.paragraphs[0]
    p.text = "AGI LiveOps"
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgbcolor('00d97e')
    p.alignment = PP_ALIGN.CENTER

    sub = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(12.3), Inches(0.8))
    tf = sub.text_frame
    p = tf.paragraphs[0]
    p.text = "让海信直播，不仅领先，而且无懈可击"
    p.font.size = Pt(24)
    p.font.color.rgb = hex_to_rgbcolor('e6edf3')
    p.alignment = PP_ALIGN.CENTER

    slogan = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.5))
    tf = slogan.text_frame
    p = tf.paragraphs[0]
    p.text = "极致效率 · 绝对掌控"
    p.font.size = Pt(18)
    p.font.color.rgb = hex_to_rgbcolor('58a6ff')
    p.alignment = PP_ALIGN.CENTER

    # 署名
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6), Inches(12.3), Inches(0.5))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.text = "方案独家提供方：聚能鼎力 AGI 实验室"
    p.font.size = Pt(14)
    p.font.color.rgb = hex_to_rgbcolor('8b949e')
    p.alignment = PP_ALIGN.CENTER

    # 保存
    output_dir = "./output"
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"海信AGI直播运营系统方案_{timestamp}.pptx"
    output_path = os.path.join(output_dir, filename)
    prs.save(output_path)

    print(f"\n✅ PPT生成成功!")
    print(f"📁 文件: {output_path}")
    print(f"📊 页数: {len(prs.slides)}")

    return output_path

if __name__ == "__main__":
    create_hisense_ppt()
