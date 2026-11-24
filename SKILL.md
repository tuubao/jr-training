---
name: jr-training
description: >-
  Fill the fixed Word training-report template from a lab guide and result
  pictures, stamp the red check on every page except the last, and save the file only on the
  user's desktop. Use this skill whenever the user wants a training report or
  lab report written, finished, or generated, including the phrases 写实训报告,
  完成实训报告, 生成实训报告, 实训报告, and 实验报告, or when they hand over a lab
  guide plus screenshots and want a Word report.
---

# Training report

Fill the fixed training-report template from the lab guide and result pictures the user provides. Save one Word file on the desktop. Do not write a second document generator, and do not hand-edit the finished docx. The script owns fonts, spacing, pictures, and the red check.

Do not guess the cover fields. A wrong name or date is printed on the report the user will submit.

The report body stays in Chinese. The form labels are Chinese, and the lab guide is Chinese. Translate instructions in this skill into action, not into an English report.

## Language when speaking to the user

Every question, and the sentence that says the file is ready, is Chinese. Do not add English words, English labels, or an English sentence around them. A file path and a date made of digits are allowed. This skill's own instructions stay in English; that English is for you, not for the user.

## Ask before doing anything

Until the cover fields are settled, do not read the lab guide, draft sections, copy the template, run the script, or create a file. The first reply is only the Chinese question below, then stop. Do not translate it, and do not put any English before or after it:

```
填写报告前，请提供下面几项。不想填的可以留空。如果都不填，直接回复“不填”。

1. 课程名称：
2. 实验名称：
3. 学生姓名：
4. 学号：
5. 日期（例如 2026-10-12）：
6. 实训室：
7. 指导教师：
```

Do not ask again when:

- The same message already gives a clear value for all six items.
- The user explicitly says to skip the cover. Treat 不填, 都不填, 跳过, 默认不填, and skip as that choice.

If only some items are given, keep those and leave the others blank. Do not ask again just to fill the gaps. Write the date at the bottom right only when a date was actually given. If the date is blank, leave the closing "year / month / day" line blank too.

When the user says to skip, leave all six cover fields blank, leave the closing date blank, and name the file `实训报告.docx`.

Do not ask for student id, major, class, or college. The template already prints the major, class, and college. The student-id cell is empty. Leave them as printed. Change one only when the user explicitly corrects it, and then put that correction in the spec. Do not append the new text after the old text.

The guide and pictures may already be attached. Read them only after the cover question is answered or explicitly skipped. If there is still no lab guide, ask only this, with no English around it: 还没有看到实训指导书，请把指导书发给我。 Do not invent an experiment. Pictures are optional. With no pictures, do not write "see the figure" and do not leave an empty figure slot.

If the guide clearly contains several different experiments, ask only this, with no English around it: 这份指导书里包含多项不同的实训，请告诉我这次要写哪一项。 Do not merge them into one report.

## What may be used as fact

The guide may be a docx, pdf, txt, md, or a photo of a paper page. Result pictures are usually png or jpg. Read all of them before writing.

Use only three kinds of facts:

- Purpose, equipment, steps, commands, parameters, and questions stated in the guide.
- Interfaces, command output, wiring, and measurements that are actually visible in the pictures.
- The cover values the user supplied.

Keep commands, IP addresses, port names, VLAN numbers, device models, pin numbers, and measured values exactly as given. Do not "rephrase" `e0/0/1` into another port, and do not turn a successful ping into a failed one. Do not add steps, numbers, or conclusions that are not in the source. If a section is missing from the guide and the pictures do not prove it, write one honest sentence for that section instead of inventing a device.

## What to change in the template

The script chooses the blank template. Do not put a template path in the spec.

1. `Template.docx` on the user's desktop, when that file exists.
2. `assets/Template.docx` shipped with this skill, when the desktop file is missing.

Both copies are read-only. The finished report is a new file on the desktop. Never save over `Template.docx`.

Keep every section heading. Fill the cells as follows:

| Place | What to do |
| --- | --- |
| Banner and the report title | Leave them |
| Course, experiment, student, date, lab room, instructor | Write the value after the colon only when one was given |
| Student id | Leave blank unless the user volunteers an id |
| Major, class, college | Keep the text already printed. Replace it only when the user explicitly corrects it |
| Purpose and requirements | Numbered points, no pictures |
| Materials and equipment | Numbered points, no pictures |
| Content, method, and steps | Numbered steps. Put process screenshots after the matching step |
| Results and discussion | Keep the existing results heading, then numbered results. Add a discussion heading only when the guide has questions or this run really had a problem. Put result screenshots here |
| Teacher comment | One comment, no more than two lines |
| Teacher signature | Always blank. Do not sign any name |
| Empty text box in the comment cell | Leave it. It is the stamp placeholder |
| Closing date at the bottom right | Fill only when a date was given. A parsed date is written as `2026年10月12日`, without the weekday |

When the cover date can be parsed, write it as `时间：2026年10月12日 星期一`. Use the weekday the user wrote. If they did not write one, use the calendar weekday. If the date cannot be parsed, copy their text after `时间：` and put the same text on the closing date line.

## How to write the body

Every section is a numbered list, not one block of prose. Do not number the points yourself. The script writes `1.` `2.` `3.` and sub-points as `（1）` `（2）`. If a point already starts with a number, the script strips it so the report does not say `1. 1.`.

Numbering restarts at 1 in each section. Purpose is usually 2 to 5 points. Each material is its own point. One step in the guide becomes one point. Do not merge unrelated steps, and do not split one step into many points. Use sub-points for small actions inside one step.

In the results section, write what was actually observed, in a way that matches the pictures. Results come first. Add discussion only when there is something to discuss.

Use complete written sentences. Do not use markdown, dash bullets, or English headings inside the report.

Write in the register of a lab report for this course, not the register of a chat. The reader is the teacher of that subject, so the wording should sound at home in that subject: a network lab uses the language of switches, ports, and reachability; a simulator lab uses the language of installation, channels, and calibration. Keep the field's own names, commands, and numbers exact, and make the sentences around them formal and specific. Name the operation, the object, and what was observed.

Do not use chatty verbs or fillers such as 搞, 弄, 整, 跑得动, 搞定, 其实, or 大概就是. A casual sentence like "先确认这台计算机跑得动模拟器" is the wrong register. The same fact in a report is "安装前核对操作系统与硬件配置，确认计算机满足模拟器的运行条件。" Formal wording is not a license to copy the guide; the 15-character limit still applies.

The script formats filled body text as SimSun (`宋体`), 12 pt, 2-character first-line indent, exactly 22 pt line spacing, left aligned. Leave the printed headings in their original formatting. Picture paragraphs must not use 22 pt exact spacing, or Word crops the picture. The script already handles that. Do not "fix" the script to give pictures the same line spacing.

## How far to rephrase

Do not copy a run of more than 15 characters from the guide. Device names, commands, addresses, and numbers do not count toward those 15 characters. The meaning stays the same.

Change word order and split long sentences. Do not change the conclusion, the object of an operation, or the level of the learning goal.

Guide: the purpose is to master basic VLAN configuration and make hosts in the same VLAN communicate.

Report point: learn how to create a VLAN on the switch and confirm that hosts in that VLAN can reach each other.

Guide: in global configuration mode, enter `vlan 10` and assign port `e0/0/1` to VLAN 10.

Report point: enter global configuration mode, create VLAN 10, and assign port `e0/0/1` to that VLAN.

A bad rewrite changes `vlan 10` to VLAN 20, turns "ping succeeded" into "can reach the internet", or says only "the experiment succeeded" without saying what succeeded.

## Where pictures go

Look at each picture and decide whether it shows the process or the outcome.

- A configuration step, command echo, or half-finished screen goes in the steps section, immediately after that step.
- A final result, connectivity test, meter reading, finished topology, or waveform goes in the results section, immediately after the result it proves.
- If a picture could be either, put an outcome in results and a procedure in steps.
- Do not put pictures in the purpose, materials, teacher comment, or cover.
- Do not place the same picture twice.
- Do not include a picture you cannot tie to the guide. Captions may describe only what is visible.

The script sets the size and centering. Width is at most 13 cm and height at most 15 cm. Landscape screenshots share one width. Do not set a width in the spec, and do not resize the file first. Write a short caption in `caption`. The script adds `图1`, `图2`, and so on across the whole report. Do not write the figure number yourself.

A caption is the label under the picture. It is not the prose between pictures. When one point has two or more pictures, a stack of pictures with only captions between them reads as a dump of screenshots. Give every one of those pictures a `lead`: one complete sentence, in the report's professional register, that says what this screen shows and what was done there. The script prints that sentence, then the picture, then the caption, then the next sentence. The reader should meet text, picture, text, picture, not a run of pictures. A single picture may omit `lead` when the point's own sentence already introduces it. Do not put the figure number in `lead`.

## Teacher comment

Write it as a teacher looking at this report, not as the student saying "I finished it". Leave the signature line blank.

Mention the real subject of this experiment. Do not reuse one sentence for every report. Cover structure, understanding of this topic, a clear process, evidence of hands-on work, and formatting. Do not give a score, do not write a name, and do not sign.

It must fit on two lines. Count each Chinese character and Chinese punctuation mark as 1, and each two letters or digits as 1. Stay at or under 70. If the script exits with code 2, shorten only the comment and run again. Do not rewrite the steps.

These show the length. Do not paste them unchanged into every report:

- 该报告结构完整，对VLAN划分认知到位，流程描述清晰，体现了较强的动手能力，格式规范，表达流畅。
- 该报告条理清楚，对静态路由的配置理解准确，步骤记录完整，能看出实际操作过程，格式规范。

## Red check on every page except the last

The final physical page must have no red check, even when it contains screenshots. A one-page report has no red check. Earlier pages each have one check, in front of text and screenshots, sized and positioned automatically for that physical page.

Use the actual page width, height, orientation and margins reported by Word/WPS after pagination, including different paper sizes in different sections. Fit the original check aspect ratio inside 60% of the writable width and 55% of the writable height, and center it in that area. Include mirrored margins and the binding gutter. Do not reuse the template's fixed check dimensions or offsets, stretch the check, or change the paper size to fit it. The check must stay inside the page's writable area.

Do not put a red check in a repeating header: that also stamps the final page. The script removes the template ink and old check-sized header/body copies, then uses `scripts/paginate_docx.ps1` to inspect a private temporary copy with Microsoft Word or WPS Writer. Actual physical pages determine the last page; section page-number restarts, text length and screenshot count do not.

Invisible temporary paragraph bookmarks identify safe anchors. The script puts one transparent check in the body on physical pages 1 through N-1, with page-relative positioning and cell clipping disabled. It measures the document again to confirm the anchors and page count have not moved, then removes the temporary bookmarks. Do not add checks manually, restore the old header picture, or hand-edit the finished document.

This step requires PowerShell on Windows and an installed Microsoft Word or WPS Writer. The script prefers PowerShell 7 when available and tries the other document application if the first cannot open the private temporary report. The helper opens that copy read-only; it does not edit the user's open document. If pagination or page dimensions are unavailable, incomplete, or change while stamping, stop with the error. Do not fall back to repeating headers or claim a finished file was saved. Keep the original template read-only.

## Page flow without large blank bands

Before measuring pages or stamping, `scripts/report_layout.py` removes the blank form's reserved minimum/fixed heights from the four filled body sections: purpose, materials, steps and results. Let these rows and body paragraphs continue naturally across pages. Explicitly disable paragraph `keepNext`, `keepLines` and `pageBreakBefore` in these cells, including their headings: WPS can propagate a table heading's `keepNext` through the body rows and push all body content off the cover page. Do not keep a whole section together or force it onto a new page. Compact only the empty paragraph separating the template tables and the mandatory empty paragraph after the last table. Keep the cover, borders, page margins, score box, comment and signature area intact.

Preview every generated page in Word/WPS. There must be no avoidable blank band like a section being pushed to the next page while several lines would still fit. Check row heights, paragraph keep settings and blank separator paragraphs before accepting that layout. Do not conceal the gap by enlarging the red check, adding filler text, or shrinking all report text. Pictures must remain readable and uncropped; if a large picture forces a substantial gap, adjust its fitting in the generator and preview again rather than hand-editing the finished report. Ordinary page margins and the intended teacher/signature space are preserved.

Confirm that the final page has no check, and check the scale and boundaries of checks on earlier text and screenshot pages. Stamping is an overlay and must not change text positions or the page count.

## Spec file

Write UTF-8 JSON in the temp directory, for example `%TEMP%\jr-training-spec.json`. Do not leave the JSON on the desktop.

```json
{
  "meta": {
    "course": "计算机网络",
    "experiment": "VLAN划分",
    "student": "张三",
    "time": "2026-10-12",
    "lab": "网络实验室",
    "teacher": "李老师"
  },
  "purpose": [
    "学会在交换机上划分 VLAN，并确认处于同一 VLAN 的主机可以互相通信。"
  ],
  "materials": [
    "二层交换机 1 台",
    "测试计算机 2 台",
    "网线若干"
  ],
  "steps": [
    {
      "text": "进入全局配置模式后创建 VLAN 10，再把端口 e0/0/1 划入这个 VLAN。",
      "children": ["创建 VLAN 10", "将端口 e0/0/1 划入 VLAN 10"],
      "images": [
        {
          "path": "C:/Users/admin/Desktop/vlan-create.png",
          "lead": "在全局配置模式中创建 VLAN 10。",
          "caption": "VLAN 10 的创建回显"
        },
        {
          "path": "C:/Users/admin/Desktop/vlan-port.png",
          "lead": "将端口 e0/0/1 划入 VLAN 10，并核对端口状态。",
          "caption": "端口划入 VLAN 10 后的状态"
        }
      ]
    }
  ],
  "results": [
    {
      "text": "PC1 向 PC2 发起 ping，能够收到应答。",
      "images": [
        {
          "path": "C:/Users/admin/Desktop/ping.png",
          "caption": "两台主机 ping 通的回显"
        }
      ]
    }
  ],
  "discussion": [
    "端口没有划入对应 VLAN 时，两台主机不在同一广播域，ping 不会通。"
  ],
  "review": "该报告结构完整，对VLAN划分认知到位，流程描述清晰，体现了较强的动手能力，格式规范，表达流畅。"
}
```

Rules:

- Use `""` or omit a meta field that has no value. Do not write placeholders such as "none".
- Do not include `student_id`, `major`, `class_name`, or `college` unless the user explicitly corrected that printed value.
- `purpose`, `materials`, `steps`, and `results` must be non-empty arrays.
- Use `[]` or omit `discussion` when there is nothing to discuss.
- A point may be a string, or an object with `text`, `children`, and `images`.
- `images[].path` is an absolute local path. `caption` cannot be empty.
- When a point has two or more pictures, every picture needs `lead`. `lead` is one full sentence printed immediately before that picture. A caption does not replace it. One picture may omit `lead`.
- Sub-points are text only. Do not hang images on them.
- `review` is required and its display width must be 70 or less.
- The VLAN story above only shows the shape of the file. Real content must come from the current guide and pictures.

## Run the script

Find the directory that contains this `SKILL.md`, then run only its script:

```text
python "<skill directory>/scripts/fill_report.py" "%TEMP%\jr-training-spec.json"
```

The script uses PowerShell plus installed Microsoft Word or WPS Writer to identify the physical last page and each page's dimensions. Keep `scripts/paginate_docx.ps1` and `scripts/report_layout.py` beside `fill_report.py`. Pagination is required even for a one-page report.

On success the script prints `TEMPLATE:` and `WROTE:`. `WROTE:` is the full path of the new file on the desktop. The file name is the experiment name. Characters that are illegal in a Windows file name are removed. If the experiment name was skipped, the name is `实训报告.docx`. An existing file is not overwritten; the script uses `_2`, `_3`, and so on. It will not overwrite `Template.docx`.

If Python says `python-docx` is missing, run `python -m pip install python-docx` and run the same command again. Do not switch libraries.

Exit code 2 means the comment is too long. Shorten `review` and run again. Any other `ERROR:` line means the spec needs a fix: a missing image, an empty caption, pictures stacked without a `lead` sentence, a picture in the wrong section, or invalid JSON. Fix the spec and run again. Do not patch the docx by hand.

After success, delete the spec file from the temp directory. The desktop should contain only the Word file.

## What to say when finished

Give the user the full `WROTE:` path and one Chinese sentence saying the report is on the desktop. Do not add an English sentence. Do not paste the report into the chat, and do not copy the file into the project or the downloads folder.

## Check before running

- The six cover fields were asked, already provided, or explicitly skipped.
- Name, instructor, lab, and date were not guessed.
- Major, class, college, and student id were not changed on your own.
- All four sections have points, and the script adds the numbers.
- Commands, addresses, models, and numbers match the guide or the pictures.
- Long sentences were not copied, and the conclusion was not changed.
- Every picture has a place and a caption, and the caption does not include `图1`.
- Two or more pictures on one point each have a `lead` sentence, so pictures are not stacked with only captions between them.
- The body uses the professional register of this lab, not chatty wording.
- The comment sounds like a teacher, names this experiment, stays within width 70, and the signature line is empty.
- The only writer is `scripts/fill_report.py`. The blank template was not overwritten.
- The red check is left to the script. Physical pages 1 through N-1 each have one check over text and pictures; the final page has none. A one-page report has no check. No repeating header check or temporary pagination bookmarks remain.
- Check sizes follow each page's actual writable dimensions and retain the image aspect ratio. No check extends into the margins or beyond the page.
- All pages were previewed: the filled body rows flow without the blank template's reserved heights or avoidable large blank bands between sections; cover and teacher/signature areas remain intact.
- The path you tell the user is the `WROTE:` path.
