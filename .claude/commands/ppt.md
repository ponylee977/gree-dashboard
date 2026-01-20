# /ppt - 智能战友PPT生成器

当用户使用 `/ppt` 命令时，执行以下操作:

## 使用方式

```
/ppt [主题] [可选参数]
```

### 示例

```
/ppt 2025年Q1业绩分析
/ppt 投资路演 --audience investors
/ppt 产品培训 --goal teach
```

## 执行步骤

1. **解析用户输入**
   - 提取主题
   - 识别受众类型 (如果指定)
   - 识别演示目标 (如果指定)

2. **运行生成器**
   ```bash
   cd $WORKSPACE/skills/ppt-generator/python
   python intelligent_ppt_strategist.py -t "[主题]" -a "[受众]" -g "[目标]"
   ```

3. **返回结果**
   - 显示生成进度
   - 返回PPT文件路径
   - 显示效果预判报告

## 参数映射

| 用户输入 | 受众类型 |
|----------|----------|
| 高管/领导/汇报 | c_suite |
| 董事会 | board |
| 投资/融资/路演 | investors |
| 客户/甲方 | clients |
| 团队/内部 | internal |
| 销售/提案 | sales |
| 培训/教学 | training |

| 用户输入 | 演示目标 |
|----------|----------|
| 汇报/报告 | report |
| 路演/融资 | pitch |
| 说服/提案 | persuade |
| 通知/告知 | inform |
| 培训/教学 | teach |

## 自动检测

如果用户未指定受众和目标，系统会:
1. 使用Claude分析主题内容
2. 自动识别最可能的受众
3. 推荐最佳演示目标
4. 选择匹配的视觉风格
