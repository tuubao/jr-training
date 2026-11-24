# jr-training

<p align="center">
  <b>把「实训指导书 + 实训图片」交给 Agent，自动生成一份完整的实训报告。</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-0.0.2-blue" alt="version">
  <img src="https://img.shields.io/badge/agents-Codex%20%C2%B7%20ClaudeCode%20%C2%B7%20Doubao%20%C2%B7%20DSH-green" alt="agents">
  <img src="https://img.shields.io/badge/license-MIT-orange" alt="license">
</p>

> 用法只有三步：**启动 Agent → 喂指导书 → 喂实训图片文件夹 → 回答问题**。

## ✨ 特性

- 严格按指导书撰写：只取标题为【实训X XXXXX实训】的章节，忽略其他实训项目。
- 每个步骤包含操作描述、匹配截图与图注、结果与简要分析。
- 报告头部信息（课程 / 姓名 / 学号 / 日期等）可自定义。
- 规范书面语，客观陈述，避免口语化。
- v0.0.2：修复红勾删不掉、末页误打红勾、固定尺寸红勾造成大片留白的问题。

## 📦 安装

下载 `jr-training.rar` 并解压，将整个 `jr-training/` 目录放入所用 Agent 的**全局 skills 目录**：

| Agent | 路径 |
| --- | --- |
| Codex | `C:\Users\<用户名>\.codex\skills\` |
| Claude Code | `C:\Users\<用户名>\.claude\skills\` |
| Doubao | `C:\Users\<用户名>\Doubao\skills\` |
| DeepSeek Harness | `C:\Users\<用户名>\.dsh\skills\` |

目录结构：

```text
skills/
└── jr-training/
    ├── SKILL.md
    ├── assets/
    ├── evals/
    └── scripts/
```

放入后**重启 Agent**，对话框上方出现本 skill 即表示安装成功。

## 🚀 使用

在对话框输入 `/jr-training`，按模板填写（路径替换为你自己的本地绝对路径）：

```text
## 任务
根据指导书文件，撰写一份完整的实训报告。

## 参考资料
- 指导书：@D:\MyStudy\指导书.docx
- ⚠️ 仅使用其中标题为【实训X XXXXX实训】的章节内容，忽略文档中其他实训项目
- 图片素材：@D:\MyStudy\实训图片（按需插入报告中）

## 基本信息（填入报告头部）
| 字段 | 值 |
| ---------- | --------------------- |
| 课程名称   | JR-Training Writting  |
| 实验名称   | What                  |
| 学生姓名   | 张三                  |
| 学号       | 202899999999          |
| 日期       | 2026-01-01            |
| 实训室     | D11-909               |
| 指导教师   | 李四                  |

## 撰写要求
1. **严格遵循指导书**：以指导书中的步骤、要求为框架，逐项完成，不遗漏任何环节。
2. **内容充实**：每个步骤需包含操作过程描述、关键截图（从图片文件夹中选取匹配的）、结果与简要分析。
3. **语言风格**：使用规范的书面语，客观陈述，避免口语化。
4. **图片使用**：从图片文件夹中选取与实验步骤对应的截图，附上简要图注（如"图1-XX操作结果"）。

## 其它需求
...
```

> ⚠️ **紧急提醒**：使用过程中若弹出选择框且文字含 "PDF"，请直接点「拒绝」，不要执行。

## ❓ FAQ

- **怎么复制提示词？** 鼠标悬停在代码框上，点右上角「复制」。
- **粘贴后不自动换行？** 不要用 `Ctrl+V` 或右键，改用 `Ctrl+Shift+V` 粘贴。

## 📄 License

MIT
