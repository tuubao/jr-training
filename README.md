# jr-training

<p align="center">
  <b>把「实训指导书 + 结果截图」交给 Agent，自动生成并排版一份完整的 Word 实训报告。</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-0.0.2-blue" alt="version">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20Word%2FWPS-green" alt="platform">
  <img src="https://img.shields.io/badge/license-MIT-orange" alt="license">
</p>

> 用法：**启动 Agent → 填写封面信息 → 喂指导书与截图 → 桌面领取 `实验名称.docx`**。

## ✨ 特性

- **固定模板填充**：保留封面、边框、评语与签名区；正文按章节自动编号，不遗漏环节。
- **图文对应**：过程截图进「步骤」，结果截图进「结果」，自动添加 `图1`、`图2` 图注；多图同点时补 `lead` 句，避免截图堆砌。
- **红勾自动化**：按 Word/WPS 实际物理分页，在第 1 页至倒数第 2 页各盖一个红勾，**末页不盖**；单页报告无红勾；不使用重复页眉。
- **版式修复**：移除空白模板的预留行高，消除大面积空白带；图片按比例自适应，不裁切。
- **只输出一份**：结果仅保存到桌面，绝不覆盖 `Template.docx`。

## 📦 安装

本仓库即 skill 根目录，克隆到所用的 Agent 全局 skills 目录即可：

```bash
git clone https://github.com/tuubao/jr-training.git <skills目录>/jr-training
```

也可直接把整个 `jr-training/` 文件夹复制到：

| Agent | 目录 |
| --- | --- |
| Codex | `C:\Users\<用户名>\.codex\skills\` |
| Claude Code | `C:\Users\<用户名>\.claude\skills\` |
| Doubao | `C:\Users\<用户名>\Doubao\skills\` |
| DeepSeek Harness | `C:\Users\<用户名>\.dsh\skills\` |

目录结构：

```text
jr-training/
├── SKILL.md                 # 技能本体（agent 指令）
├── README.md
├── assets/
│   ├── Template.docx        # 固定报告模板（只读）
│   └── red-check.png        # 红勾素材
├── evals/
│   └── evals.json           # 测试用例
└── scripts/
    ├── fill_report.py       # 主生成器
    ├── paginate_docx.ps1    # 物理分页探测
    ├── report_layout.py     # 版式修复
    ├── test_red_checks.py
    └── test_report_layout.py
```

依赖：**Windows + PowerShell + Microsoft Word 或 WPS Writer**；Python 需 `python-docx`（缺失时按脚本提示安装）。

## 🚀 使用

在对话框输入 `/jr-training`，然后：

1. 按提示提供课程名称、实验名称、学生姓名、学号、日期、实训室、指导教师（不想填可留空或回复「不填」）。
2. 提供实训指导书（docx / pdf / txt / md / 图片均可）和结果截图（可选）。
3. Agent 生成 spec 并调用 `scripts/fill_report.py`，报告自动落到桌面。

> ⚠️ 若过程中出现含 "PDF" 字样的选择框，请直接点「拒绝」，不要执行。

## ❓ FAQ

- **红勾为什么这么处理？** 重复页眉会连末页一起盖章；改用物理分页定位，末页无勾、不拉伸、不越界。
- **正文为什么这么"啰嗦"？** 需符合该课程的书面语体，同时连续摘抄不超过 15 字、不改变结论。
- **粘贴提示词不换行？** 用 `Ctrl+Shift+V`，不要用 `Ctrl+V` 或右键粘贴。

## 📄 License

MIT
