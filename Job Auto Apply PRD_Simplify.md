# Job Auto Apply PRD_Simplify

- 文档版本：v0.1
- 创建时间：2026-08-30 14:24:20
- 项目名称：Case A_Job Auto Apply
- 项目根目录：`C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply`
- 运行环境：Hermes Agent Desktop
- 演示范围：仅英文简历 + 海外/英文官网招聘申请页面
- 核心工具：Hermes browser automation + autoapply-mcp
- 明确不使用：牛客网申助手、Simplify Copilot、其他浏览器插件
- 核心边界：半自动填表，不自动最终提交

---

## 1. 背景与目标

本 PRD 是为案例演示场景简化后的半自动英文简历投递工作流需求文档。

原完整方案覆盖中文网站、英文网站、浏览器插件辅助、定时状态检查等多场景；本简化版本只保留适合演示的核心路径：

```text
英文简历 → 英文官网岗位搜索 → 岗位 Excel 整理 → 用户确认 → Hermes/autoapply-mcp 半自动填写 → 用户手动提交 → Excel 记录
```

目标是展示 Hermes Agent 如何作为主控，在本地项目目录中完成：

1. 读取英文简历并建立结构化英文 profile；
2. 根据用户目标行业和岗位大类搜索正在开放的英文官网岗位；
3. 将公司、岗位、JD、投递链接整理到本地 Excel；
4. 用户确认要投递的岗位；
5. Agent 使用 Hermes browser automation 和/或 autoapply-mcp 辅助填写申请表；
6. 在最终提交前停止，由用户人工检查和提交；
7. 根据用户确认，将已投递岗位写入 Excel 记录。

---

## 2. 项目配置需求

### 2.1 VSCode-like Project Workspace

本项目要求仿照 VSCode workspace 的使用方式：

- 所有项目相关代码、配置、Excel、日志、profile、MCP 配置文件、生成文件都集中在项目根目录下；
- Agent 后续写入文件时，默认以该目录作为工作路径；
- 不应将项目文件散落在用户 home、临时目录或 Hermes attachment 目录中，除非外部工具强制要求；
- 如果外部工具默认使用 `~/.autoapply`，应优先通过环境变量改为项目内 `.autoapply/`。

项目根目录固定为：

```text
C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply
```

### 2.2 推荐目录结构

```text
Case A_Job Auto Apply/
  Job Auto Apply PRD_Simplify.md
  Job Auto Apply PRD.md

  resumes/
    original/
      Resume.pdf
    pdf/
      Resume_EN.pdf

  profiles/
    resume_data.json
    profile_en.json
    application_private.example.json
    application_preferences.json

  excel/
    岗位搜索与简历投递.xlsx

  applications/
    generated_answers/
    application_packets/
    screenshots/

  logs/
    autofill_runs.md
    job_search_runs.md
    application_status_checks.md

  mcp/
    autoapply-mcp/

  .autoapply/
    profile.json
    campaign.json
    companies.json
    resumes/
    data/
    artifacts/

  scripts/
    convert_resume_profile_en.py
    update_job_excel.py
    status_check.py
```

说明：

- `profiles/profile_en.json`：Hermes-native 英文申请 profile；
- `.autoapply/profile.json`：autoapply-mcp 使用的候选人 profile；
- `excel/岗位搜索与简历投递.xlsx`：主工作簿；
- `logs/`：保存搜索、填表、状态检查记录；
- `applications/`：保存每个岗位的申请草稿、截图、packet 和审计文件。

---

## 3. 工具范围与集成边界

### 3.1 Hermes Browser Automation

Hermes browser automation 是本演示工作流的第一优先级工具。

用途：

- 打开公司官网 career 页面；
- 打开 ATS 申请页面；
- 读取岗位 JD；
- 识别申请表字段；
- 自动填写非敏感字段；
- 上传英文简历；
- 对开放题生成英文草稿；
- 对敏感字段暂停并询问用户；
- 在最终提交按钮前停止。

适用平台：

- 公司官网 career 页面；
- Greenhouse；
- Lever；
- Ashby；
- Workday 的部分官网申请流程；
- 其他英文官网申请页面。

限制：

- 不处理账号密码；
- 不破解 CAPTCHA；
- 不点击最终 Submit / Apply / Submit Application；
- 登录、验证码、最终提交由用户处理。

### 3.2 autoapply-mcp

`autoapply-mcp` 作为第二个核心工具，用于英文官网 ATS 专项流程。

定位：

```text
Greenhouse / Lever / Ashby 的岗位发现、筛选、评分、申请 packet 准备和 assisted filling 工具
```

用途：

- 读取公开 ATS board；
- 发现开放岗位；
- 按 candidate profile 和 campaign 过滤岗位；
- 对岗位进行匹配评分；
- 生成申请 packet；
- 支持 manual / assisted 模式；
- 保存审计记录。

配置原则：

```text
AUTOAPPLY_HOME = C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply\.autoapply
```

安全限制：

- 初期只允许 `manual` 或 `assisted` 模式；
- 不启用 `auto` 或自动提交；
- 不使用 LinkedIn、Indeed、Wellfound 自动化；
- 不将 Workday 作为 autoapply-mcp 的自动投递目标；
- 遇到 CAPTCHA 立即停止；
- 敏感问题必须交给用户确认。

### 3.3 明确不使用的工具

为适配案例演示，本版本不安装、不接入、不演示以下工具：

```text
牛客网申助手
Simplify Copilot
OfferLink
其他中文网申插件
LinkedIn 非官方 MCP/API
LazyApply / Sonara / AIApply 等自动海投工具
```

原因：

- 本次演示只聚焦英文简历投递官网招聘；
- 不演示中文网站；
- 不依赖第三方浏览器插件；
- 避免插件账号、权限、隐私和 UI 不稳定因素；
- 让演示路径更清晰：Hermes + autoapply-mcp。

---

## 4. 语言、地区与岗位范围

### 4.1 语言范围

本演示只使用英文流程：

```text
profile_en.json
英文简历 PDF/DOCX
英文官网申请页面
英文开放题草稿
```

不演示：

```text
中文简历
中文官网
中文校招系统
BOSS 直聘
牛客网申
中文插件辅助
```

### 4.2 地区确认

工作流启动时，Agent 仍需向用户确认目标投递国家/地区。

可选地区包括：

- Singapore；
- Hong Kong；
- China / APAC 英文岗位；
- Remote；
- US；
- Europe；
- 其他用户指定地区。

### 4.3 岗位范围

默认面向英文官网中的以下岗位大类：

- Data Analyst；
- Business Analyst；
- Data Scientist；
- Machine Learning / AI；
- Quantitative / FinTech；
- Technical Consultant；
- Analytics / Strategy / Operations Analytics。

Agent 推荐岗位时应综合：

- 专业背景匹配度；
- Python / SQL / Statistics / ML 技能匹配；
- 金融、咨询、数据分析实习经历匹配；
- 地区与工作授权风险；
- 年限要求；
- 是否接受应届生或 early career；
- 岗位 JD 的真实要求。

---

## 5. Token 节约原则

案例演示应尽量节约 token，并优先使用结构化、可复用的数据。

原则：

1. 用户已有公司或岗位清单时，不做广泛搜索；
2. 岗位搜索优先使用 ATS API、公司 career 页面、autoapply-mcp，而不是读取大量网页全文；
3. 只提取 JD 中与岗位匹配、硬性要求、申请问题相关的内容；
4. Excel 采用增量写入，不反复全文重写；
5. 已搜索过的公司只检查新开放岗位；
6. 开放题只在用户确认要投递该岗位后生成；
7. 对长 JD 只保留摘要、关键要求和原文链接；
8. 使用 autoapply-mcp 处理 Greenhouse / Lever / Ashby 时，优先复用其发现和评分结果；
9. 不为中文流程、浏览器插件、LinkedIn 非官方自动化消耗演示 token。

---

## 6. 主 Excel 工作簿需求

### 6.1 工作簿名

在本地路径生成 Excel 文件：

```text
C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply\excel\岗位搜索与简历投递.xlsx
```

工作簿名称：

```text
岗位搜索与简历投递
```

### 6.2 Sheet 1：意向投递公司

Sheet 名称：

```text
意向投递公司
```

用途：

- 保存用户想投递的英文官网目标公司；
- 保存 Agent 根据英文简历和目标岗位大类补充的公司；
- 保存公司 career 页面和 ATS board 链接；
- 作为后续岗位搜索来源表。

建议字段：

| 字段名 | 说明 |
|---|---|
| company_id | 公司唯一 ID，由 Agent 生成 |
| Company Name | 英文公司名 |
| 国家/地区 | 目标地区 |
| 城市 | 分公司或岗位所在城市 |
| 是否目标地区有分公司 | 是/否/待确认 |
| 行业 | FinTech / Consulting / Internet / Data Services 等 |
| 岗位大类 | Data Analyst / Data Scientist / ML / BA 等 |
| 匹配简历 | 默认 `profile_en.json` |
| Career 官网链接 | 公司官网 career 页面 |
| ATS 类型 | Greenhouse / Lever / Ashby / Workday / 自建 / 未知 |
| ATS Board 链接 | 若可识别，填写 ATS board 链接 |
| LinkedIn 公司岗位链接 | 可选，仅作为岗位线索，不做自动化投递 |
| 信息来源 | 用户提供 / Agent 搜索 / autoapply-mcp |
| 最近检查时间 | 最近一次检查该公司岗位时间 |
| 备注 | 其他说明 |

规则：

- 用户可直接粘贴公司清单到该 sheet；
- Agent 只补全缺失列和校验链接；
- 如需覆盖用户已有内容，必须先获得用户确认；
- 不需要维护 BOSS、牛客、中文校招链接。

### 6.3 Sheet 2：开放岗位列表

Sheet 名称：

```text
开放岗位列表
```

用途：

- 保存英文官网或 ATS 中发现的开放岗位；
- 保存 JD 摘要、JD 链接、投递链接；
- 支持增量更新。

建议字段：

| 字段名 | 说明 |
|---|---|
| job_id | 岗位唯一 ID，优先 ATS job id，否则 URL hash |
| company_id | 对应公司 ID |
| Company Name | 公司名 |
| Job Title | 岗位名称 |
| Role Category | 岗位大类 |
| Country/Region | 国家/地区 |
| City | 城市 |
| Work Mode | onsite / hybrid / remote / unknown |
| Language | 固定为 English |
| JD Summary | 精简 JD 摘要 |
| Key Requirements | 关键要求 |
| JD URL | Job description 页面链接 |
| Application URL | 投递链接 |
| Source Platform | Company Site / Greenhouse / Lever / Ashby / Workday / Other |
| Match Score | 0-100 |
| Match Rationale | 匹配理由 |
| Risk / Gaps | 年限不符、签证未知、地点不符等 |
| Discovered At | 首次发现时间 |
| Last Checked At | 最近确认开放时间 |
| Is Applied | Yes / No / Pending |
| User Decision | Apply / Skip / Pending |
| Notes | 备注 |

规则：

- Agent 在用户要求时顺序检查“意向投递公司”；
- 如果已查找过一轮，后续只关注新增岗位；
- 新岗位需要与现有 `job_id` 或 URL 去重；
- Agent 在对话框中展示候选岗位 JD 摘要和投递链接；
- 用户确认后才启动半自动填写。

### 6.4 Sheet 3：已投递记录

Sheet 名称：

```text
已投递记录
```

用途：

- 记录用户实际手动提交过的岗位；
- 记录 JD 链接、投递链接、投递时间和状态变化。

建议字段：

| 字段名 | 说明 |
|---|---|
| application_id | 投递记录唯一 ID |
| job_id | 对应开放岗位 job_id |
| Company Name | 公司名 |
| Job Title | 岗位名称 |
| Country/Region | 地区 |
| City | 城市 |
| Resume/Profile Used | 默认英文简历 / `profile_en.json` |
| JD URL | JD 链接 |
| Application URL | 投递链接 |
| Platform | Company Site / Greenhouse / Lever / Ashby / Workday / Other |
| Applied At | 用户确认已提交的时间 |
| Tool Combination | Hermes browser automation / autoapply-mcp / both |
| User Submitted Final Application | Yes / No / Pending |
| Current Status | submitted / under review / assessment / interview / rejected / offer / unknown |
| Status Updated At | 最近状态更新时间 |
| Status 1 | 第一次状态变化 |
| Updated At 1 | 第一次状态变化时间 |
| Status 2 | 第二次状态变化 |
| Updated At 2 | 第二次状态变化时间 |
| Status 3 | 第三次状态变化 |
| Updated At 3 | 第三次状态变化时间 |
| Status 4 | 第四次状态变化 |
| Updated At 4 | 第四次状态变化时间 |
| Notes | 备注 |

规则：

- 只有用户明确表示已经手动提交，或要求记录为已投递，才写入该 sheet；
- 如果只是填到最终提交前，应记录为 `filled_not_submitted`，不应伪装为已投递；
- 状态更新采用追加列方式，不覆盖历史状态；
- 没有检测到状态变化时，不改动 Excel。

---

## 7. 英文简历与岗位大类关联需求

### 7.1 用户输入

用户可能提供：

- 一份英文简历；
- 一个或多个英文岗位大类；
- 目标国家/地区；
- 已调研公司列表；
- 已调研岗位 Excel。

### 7.2 Agent 处理规则

Agent 应建立：

```text
英文简历 / profile_en.json ↔ 行业 ↔ 岗位大类 ↔ 公司类型 ↔ 地区
```

的关联关系。

建议生成文件：

```text
profiles/resume_role_mapping.json
```

示例结构：

```json
{
  "mappings": [
    {
      "profile": "profile_en.json",
      "resume_file": "resumes/pdf/Resume_EN.pdf",
      "target_regions": ["Singapore", "APAC", "Remote"],
      "industries": ["FinTech", "Consulting", "Internet", "Data Services"],
      "role_categories": ["Data Analyst", "Data Scientist", "Business Analyst", "Machine Learning"]
    }
  ]
}
```

### 7.3 匹配原则

Agent 推荐公司和岗位时，应综合：

- 教育背景；
- 技术栈；
- 数据/统计/机器学习能力；
- 实习经历；
- 项目和竞赛经历；
- 地理位置；
- 英文岗位语言要求；
- 年限要求；
- 签证/工作授权风险；
- 用户明确排除的行业或公司。

---

## 8. 岗位搜索流程需求

### 8.1 可选入口 A：Agent 主动搜索

当用户只提供英文简历和目标行业/岗位大类时，Agent 应：

1. 向用户确认目标国家/地区；
2. 解析英文简历并建立 `profile_en.json`；
3. 按地区、行业和岗位大类搜索合适公司；
4. 检查公司是否在目标地区有分公司或开放岗位；
5. 补充 career 页面、ATS board 链接和可选 LinkedIn 公司岗位链接；
6. 写入 Excel 的“意向投递公司”sheet；
7. 在用户要求时继续搜索开放岗位并写入“开放岗位列表”。

### 8.2 可选入口 B：用户上传或粘贴公司/岗位 Excel

当用户已经调研过并上传公司或岗位列表时，Agent 应：

1. 读取用户提供的 Excel；
2. 判断它更接近“意向投递公司”还是“开放岗位列表”；
3. 映射到主工作簿 schema；
4. 只补全缺失信息，不随意覆盖用户原始内容；
5. 如需修改用户已有内容，先说明差异并请求确认。

### 8.3 岗位增量搜索

当用户要求查找岗位，或定时任务触发时：

1. 顺序读取“意向投递公司”；
2. 对每家公司检查 career 页面和 ATS board；
3. 对 Greenhouse / Lever / Ashby 优先使用 autoapply-mcp 或 ATS API；
4. 找出符合用户行业与岗位大类的新开放岗位；
5. 与“开放岗位列表”中已有 `job_id` 或 URL 去重；
6. 只新增新岗位，或在用户确认后更新状态变化；
7. 在对话框中汇报新增岗位 JD 摘要和投递链接；
8. 等用户确认后才启动申请填写。

---

## 9. 半自动英文官网投递流程需求

### 9.1 启动条件

半自动投递流程只能在以下条件满足时启动：

1. 用户明确选择要投递的岗位；
2. 已确定使用英文简历 / `profile_en.json`；
3. 已确认投递平台和工具组合；
4. 敏感字段缺口已列出；
5. 用户知道 Agent 不会点击最终提交。

### 9.2 工具组合选择

默认组合：

```text
Hermes browser automation
```

ATS 专项组合：

```text
autoapply-mcp + Hermes browser automation
```

选择规则：

| 平台 | 推荐工具组合 |
|---|---|
| Greenhouse | autoapply-mcp + Hermes browser automation |
| Lever | autoapply-mcp + Hermes browser automation |
| Ashby | autoapply-mcp + Hermes browser automation |
| Workday | Hermes browser automation only |
| 自建官网申请页 | Hermes browser automation only |
| LinkedIn | 不作为本演示核心；如出现，仅作为岗位线索，不自动投递 |

### 9.3 自动填写范围

允许自动填写：

- First Name；
- Last Name；
- Email；
- Phone，但若号码存在冲突需用户确认；
- Education；
- Work Experience；
- Projects；
- Skills；
- Languages；
- Resume Upload；
- 基于真实经历生成的英文开放题草稿。

不允许自动猜测：

- Work authorization；
- Visa status；
- Sponsorship requirement；
- Expected salary；
- Current location；
- Available start date；
- Legal attestations；
- Criminal history；
- EEO / disability / veteran / demographic 信息；
- 背景调查授权；
- 任何真实性承诺。

### 9.4 提交边界

硬性规则：

```text
Agent 永远不自动点击最终 Submit / Apply / Submit Application 按钮。
```

Agent 必须在最终提交前：

1. 汇报已填写字段；
2. 汇报未填写或需确认字段；
3. 汇报开放题草稿来源；
4. 提醒用户自行检查；
5. 等待用户自行提交或放弃；
6. 只有用户确认已提交后，才写入“已投递记录”。

---

## 10. 定时任务需求

### 10.1 每日岗位搜索 Timer

用户希望设置 timer：

```text
当地时间每天早晨 10:00
```

触发任务：

1. 读取“意向投递公司”sheet；
2. 检查每家公司是否有新开放英文官网岗位；
3. 只关注用户目标行业和岗位大类；
4. 如果已有历史搜索记录，只汇报新开放岗位；
5. 将新岗位写入“开放岗位列表”；
6. 在对话框中给出新岗位 JD 摘要和投递链接；
7. 若用户接受某岗位，再启动半自动投递流程。

### 10.2 每日投递状态检查 Timer

同一 timer 或另一个 timer 在每天 10:00 触发：

1. 读取“已投递记录”sheet；
2. 检查已投递岗位是否有状态更新；
3. 可能状态来源包括：
   - 申请门户；
   - ATS 页面；
   - 用户授权后的邮件状态通知；
   - 用户手动提供的状态更新；
4. 如果有状态变化，则写入新的 `Status N / Updated At N`；
5. 如果没有状态变化，不改动 Excel；
6. 如果需要登录或验证码，暂停并请求用户介入。

### 10.3 Timer 安全规则

- Timer 不自动投递；
- Timer 不自动登录；
- Timer 不自动填写申请表；
- Timer 不删除 Excel 中任何数据；
- Timer 的输出必须可继续交互，让用户确认是否投递；
- Timer prompt 必须自包含，不依赖当前聊天上下文。

---

## 11. Excel 数据修改与权限规则

### 11.1 禁止私自删除

未经用户明确允许，Agent 不得删除 Excel 内任何信息。

包括：

- 不删除公司；
- 不删除岗位；
- 不删除历史状态；
- 不删除用户备注；
- 不删除旧链接；
- 不删除手动粘贴内容。

### 11.2 更新规则

Agent 可以在以下情况下新增信息：

- 用户要求搜索并补全；
- 定时任务发现新开放岗位；
- 用户确认已投递；
- 用户确认状态变化；
- 自动检查发现状态更新且证据明确。

Agent 在以下情况下需要用户确认后才能修改：

- 覆盖已有字段；
- 标记公司无效；
- 标记岗位关闭；
- 合并重复公司/岗位；
- 删除重复项；
- 修正用户手动填写的信息。

### 11.3 审计记录

每次写入 Excel 时，应在日志中记录：

- 时间；
- 操作类型；
- 修改 sheet；
- 新增/更新的行数；
- 数据来源；
- 是否经用户确认；
- 相关岗位链接。

---

## 12. 用户交互需求

### 12.1 初始确认

Agent 在启动演示流程前应确认：

- 目标国家/地区；
- 使用哪份英文简历；
- 目标行业；
- 目标岗位大类；
- 是否有不投公司/行业；
- 是否已有公司/岗位 Excel；
- 是否允许创建或更新 `岗位搜索与简历投递.xlsx`。

### 12.2 投递前确认

Agent 在投递前应确认：

- 用户选择哪些岗位投递；
- 使用哪份英文简历 / profile；
- 使用 Hermes browser automation 还是 autoapply-mcp；
- 电话号码选择；
- 当前所在地；
- work authorization / visa sponsorship；
- expected salary / available start date；
- 开放题草稿是否可用。

### 12.3 投递后确认

Agent 在投递后应确认：

- 用户是否已经手动提交；
- 是否记录为已投递；
- 是否需要记录备注；
- 是否需要安排后续状态检查。

---

## 13. 安全与合规要求

硬性要求：

1. 不自动点击最终提交；
2. 不处理账号密码；
3. 不破解 CAPTCHA；
4. 不使用代理池、指纹伪装或反检测工具；
5. 不调用 LinkedIn 非官方 API；
6. 不伪造经历、技能、教育背景；
7. 不自动回答敏感字段；
8. 不删除 Excel 历史信息；
9. 不未经确认覆盖用户手动输入；
10. 所有申请内容必须基于英文简历或用户确认材料；
11. `autoapply-mcp` 只能在 manual / assisted 范围内使用，不能自动提交。

---

## 14. MVP 范围

### 14.1 MVP 必须实现

1. 在项目目录下生成主 Excel：`岗位搜索与简历投递.xlsx`；
2. 建立至少三个 sheet：
   - `意向投递公司`；
   - `开放岗位列表`；
   - `已投递记录`；
3. 从英文简历生成或维护：
   - `profile_en.json`；
   - `.autoapply/profile.json`；
4. 支持用户手动提供公司/岗位 Excel 并映射进主工作簿；
5. 支持按公司查找开放英文官网岗位；
6. 支持 Greenhouse / Lever / Ashby 通过 autoapply-mcp 辅助发现与准备申请；
7. 支持用户确认岗位后启动半自动英文填表；
8. 支持填写前敏感字段清单；
9. 支持在最终提交前停止；
10. 支持用户确认后记录已投递岗位；
11. 支持每日 10:00 的岗位更新和状态检查 timer 设计。

### 14.2 MVP 暂不实现

1. 不实现全自动提交；
2. 不实现中文简历投递演示；
3. 不实现中文官网或中文校招流程；
4. 不安装或使用牛客网申助手；
5. 不安装或使用 Simplify Copilot；
6. 不安装或使用 OfferLink；
7. 不实现 LinkedIn 非官方 API；
8. 不实现验证码破解；
9. 不实现自动批量海投；
10. 不强制接入第三方 CRM。

---

## 15. 验收标准

### 15.1 文件与路径

- 所有项目文件均在 `C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply` 下；
- Excel、profiles、logs、applications、mcp 目录结构清晰；
- 不将项目数据散落到无关目录。

### 15.2 Excel 工作簿

- 存在 `岗位搜索与简历投递.xlsx`；
- 至少包含 `意向投递公司`、`开放岗位列表`、`已投递记录`；
- 字段满足英文官网岗位搜索和投递记录需求；
- 新增和状态更新可追溯；
- 不经用户允许不会删除历史数据。

### 15.3 工作流

- Agent 能根据英文简历和目标岗位大类整理公司；
- Agent 能检查英文官网或 ATS 新开放岗位；
- Agent 能给出 JD 摘要和投递链接；
- Agent 能在用户确认后顺序启动半自动英文投递；
- Agent 能根据平台选择 Hermes browser automation 或 autoapply-mcp；
- Agent 能在最终提交前停止；
- Agent 能在用户确认后记录投递结果。

### 15.4 安全

- 敏感字段不会被猜测；
- 登录和密码由用户自己处理；
- 不自动提交；
- 不绕过平台限制；
- 申请内容不编造；
- 不依赖浏览器插件完成演示。

---

## 16. 后续开发建议

建议按以下顺序开发演示版：

1. 整理项目目录结构；
2. 创建 `excel/岗位搜索与简历投递.xlsx` 模板；
3. 从英文简历生成 `profiles/profile_en.json`；
4. 生成 `.autoapply/profile.json` 所需字段映射；
5. 建立 `application_preferences.json` 和敏感字段清单；
6. 实现英文官网岗位搜索与 Excel 更新流程；
7. 选择一个 Greenhouse / Lever / Ashby 岗位做 autoapply-mcp 准备流程；
8. 使用 Hermes browser automation 打开真实官网申请页并半自动填写；
9. 在最终提交前停止并让用户检查；
10. 用户确认提交后写入“已投递记录”；
11. 设计每日 10:00 timer，但在演示阶段可先手动触发测试。

---

## 17. 当前明确不做的事项

- 不安装牛客网申助手；
- 不安装 Simplify Copilot；
- 不演示中文简历投递；
- 不演示中文官网/校招系统；
- 不点击任何最终提交按钮；
- 不处理账号密码；
- 不破解验证码；
- 不进行自动海投。

本 PRD 仅用于案例演示版 Job Auto Apply 工作流的后续开发依据。
