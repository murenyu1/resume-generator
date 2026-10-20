# Resume Generator (简历专家) 🚀

> 专为 DeepSeek Harness (DSH) 打造的专业级简历架构师与生成器 Skill。  
> 帮助求职者将零散背景经历，重构成兼具 **STAR 原则**、**Google XYZ 法则** 与 **高通过率 ATS 友好度** 的高水准简历。

---

## ✨ 核心特性

- 🎯 **真实性红线**：严禁造假，在真实经历的边界内最大化个人价值与亮点。
- ⚡ **STAR / Google XYZ 深度重构**：杜绝“负责日常开发”等流水账，全面采用强动词与度量指标定级。
- 🏢 **目标岗位 (JD) 靶向对齐**：智能提取 JD 硬技能与业务场景关键词，进行倒金字塔视觉重排。
- 📄 **多模态与全格式交付**：
  - **优雅 Markdown**：中英文排版规范，便于即时阅读与微调。
  - **商业级 Word (.docx)**：基于 Python 原生构建标准 A4 窄边距、深蓝商务配色与双端无框对齐表格。
  - **打印级单页 HTML 模板**：支持在浏览器中一键通过 `Ctrl+P` 无损“另存为 PDF”，杜绝排版错位和空白悬垂页。

---

## 📂 目录架构

```text
resume-generator/
├── SKILL.md                               # Skill 核心入口与工作流编排
├── README.md                              # 项目说明文档
└── references/
    ├── bullet-formulas.md                 # 经历与 Bullet 黄金公式库及强动词词典
    ├── ats-tailoring.md                   # ATS 筛选避坑与 JD 靶向匹配指南
    ├── docx-generator-template.py         # 原生 python-docx 商业级 Word 生成脚本模板
    └── single-page-resume.html            # 打印级现代单页 HTML 简历模板
```

---

## 🛠️ 安装与使用 (DSH 环境)

### 1. 安装到 DSH 本地 Skill 目录
将本项目克隆或放置到 DSH 的 skills 目录下：
```bash
# 默认路径
~/.dsh/skills/resume-generator
```

### 2. 使用方法
在 DeepSeek Harness 中与 AI 对话时，直接发送您的个人经历或输入：
> “使用 resume-generator 帮我写一份简历，我的求职意向是资深后端架构师，以下是我的经历……”

AI 将自动加载本技能，按阶段进行资料诊断、STAR 提炼、JD 靶向匹配，并输出对应格式的简历。

---

## 📄 开源许可证

本项目基于 [MIT License](LICENSE) 开源。
