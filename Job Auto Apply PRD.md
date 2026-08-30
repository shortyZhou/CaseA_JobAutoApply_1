# Case A_Job Auto Apply — 半自动简历投递工作流需求文档

- 文档版本：v0.1
- 创建时间：2026-08-30 12:52:08
- 项目名称：Case A_Job Auto Apply
- 项目根目录：`C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply`
- 运行环境：Hermes Agent Desktop
- 目标：构建一个可审计、可暂停、可人工确认的中英文半自动简历投递工作流

---

## 1. 背景与目标

用户希望在 Hermes Agent 中开发一套用于求职申请的半自动工作流，覆盖中文网站和英文网站的岗位搜索、岗位整理、简历匹配、申请辅助填写、投递记录维护和定时岗位更新。

该工作流的核心原则不是全自动海投，而是：

```text
Search → Filter → Fill → Confirm → Record
```

即：

1. Agent 帮助用户寻找符合行业、岗位大类、地区要求的公司和开放岗位；
2. 将公司、岗位、JD、投递链接整理到本地 Excel；
3. 用户确认要投递的岗位；
4. Agent 使用 Hermes browser automation / autoapply-mcp / 浏览器插件辅助填写申请表；
5. Agent 在最终提交前停止，由用户人工检查并决定是否提交；
6. Agent 将已投递记录和状态更新写回本地 Excel。

---

## 2. 项目配置需求

### 2.1 VSCode-like Project Workspace

本项目要求完全仿照 VSCode 的 workspace/project 使用方式：

- 所有项目相关代码、配置、Excel、日志、profile、MCP 配置文件、生成文件都集中在项目根目录下；
- Agent 后续写入文件时，默认以该目录作为工作路径；
- 不应将项目文件散落在用户 home、临时目录或 Hermes attachment 目录中，除非工具强制要求；
- 若外部工具默认使用其他目录，例如 `~/.autoapply`，应优先通过环境变量或配置改为项目内目录。

项目根目录固定为：

```text
C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply
```

### 2.2 推荐目录结构

```text
Case A_Job Auto Apply/
  README.md
  半自动简历投递工作流需求文档.md
  job_autofill_session_record.md

  resumes/
    original/
      Zhou Tianyi_Resume2023_draft.docx
      周天怡_简历2023_草稿.docx
    pdf/
      Zhou_Tianyi_Resume_EN.pdf
      Zhou_Tianyi_Resume_ZH.pdf

  profiles/
    resume_data.json
    profile_en.json
    profile_zh.json
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
    convert_resume_profiles.py
    update_job_excel.py
    status_check.py
```

说明：

- `profiles/`：保存 Hermes-native 工作流使用的结构化数据；
- `.autoapply/`：保存 autoapply-mcp 的运行数据；
- `excel/岗位搜索与简历投递.xlsx`：主工作簿；
- `logs/`：保存每次搜索、填表、状态检查的可追溯记录；
- `applications/`：保存每个岗位生成的申请材料、截图、草稿和审计文件。

---

## 3. 使用工具与集成边界

### 3.1 Hermes Browser Automation

Hermes browser automation 是本项目第一优先级工具。

用途：

- 打开岗位页面；
- 读取网页内容和岗位描述；
- 识别申请表字段；
- 填写非敏感字段；
- 上传简历；
- 生成或填写开放题草稿；
- 对敏感字段暂停并询问用户；
- 在最终提交按钮前停止。

优势：

- 无需额外安装；
- 覆盖中文和英文网页；
- 可用于 LinkedIn、企业官网、ATS、中文校招系统；
- 便于用户可视化检查。

限制：

- 登录需要用户自己完成；
- 不处理验证码破解；
- 不点击最终提交按钮；
- 对高度动态或受保护页面，可能需要人工辅助。

### 3.2 autoapply-mcp

用户计划安装 `autoapply-mcp` 作为 Agent 工作流工具。

定位：

```text
Greenhouse / Lever / Ashby 专项的岗位发现、筛选、评分、申请材料准备和 assisted filling 工具
```

用途：

- 读取公开 ATS board；
- 发现岗位；
- 按用户 profile 和 campaign 筛选岗位；
- 对岗位进行评分；
- 准备申请 packet；
- 在 manual 或 assisted 模式下协助填写；
- 保存审计记录。

限制：

- 不作为 LinkedIn 自动化工具；
- 不作为 Workday 全自动投递工具；
- 不处理中文校招系统；
- 不启用 auto-submit；
- 初期只允许 `manual` 或 `assisted` 模式。

建议配置原则：

```text
AUTOAPPLY_HOME = C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply\.autoapply
```

### 3.3 牛客网申助手

牛客网申助手由用户人工下载安装到浏览器中。

定位：

```text
中文官网 / 中文校招 / 国内网申辅助填表插件
```

Agent 职责：

- 指导用户从官方渠道下载和安装；
- 指导用户导入或维护简历信息；
- 指导用户在中文网申页面使用插件；
- 检查插件填充结果；
- 识别插件未填或误填字段；
- 对敏感字段提示人工确认；
- 记录投递流程。

边界：

- Hermes 不直接接入牛客插件 API；
- Hermes 不代替用户登录牛客账号；
- Hermes 不应让插件自动填写身份证号、政治面貌、期望薪资、亲属关系、真实性承诺等敏感字段；
- 用户手动最终提交。

### 3.4 Simplify Copilot

Simplify Copilot 由用户人工下载安装到浏览器中。

定位：

```text
英文 ATS / LinkedIn / 海外企业官网辅助填表插件
```

Agent 职责：

- 指导用户安装；
- 指导用户维护英文 profile；
- 在英文申请页面中辅助用户使用 Simplify；
- 检查其填充结果；
- 对 LinkedIn、work authorization、visa sponsorship、expected salary、EEO 等字段暂停；
- 记录投递流程。

边界：

- Hermes 不直接接入 Simplify API；
- Hermes 不代替用户登录 Simplify 或 LinkedIn；
- Simplify 仅作为辅助，不作为最终判断和提交工具。

---

## 4. 语言与地区策略

### 4.1 地区确认

工作流启动时，Agent 必须先向用户确认目标投递地区。

一级分类：

1. 中国境内地区；
2. 非中国地区。

可扩展细分：

- 中国大陆；
- 香港；
- 新加坡；
- 美国；
- 欧洲；
- Remote；
- APAC；
- 其他。

### 4.2 中文投递策略

适用：

- 中国境内官网；
- 中文校招系统；
- 国内招聘平台；
- BOSS 直聘相关岗位入口；
- 中文企业 career 页面。

默认工具组合：

```text
Hermes browser automation + 牛客网申助手
```

可选工具：

```text
OfferLink / 其他中文插件
```

默认 profile：

```text
profiles/profile_zh.json
```

### 4.3 英文投递策略

适用：

- 英文企业官网；
- LinkedIn；
- Greenhouse；
- Lever；
- Ashby；
- Workday 部分场景。

默认工具组合：

```text
Hermes browser automation + Simplify Copilot
```

ATS 专项可选：

```text
autoapply-mcp + Hermes browser automation
```

默认 profile：

```text
profiles/profile_en.json
```

### 4.4 Token 节约原则

工作流应尽量减少不必要 token 消耗。

原则：

1. 用户已经提供 Excel 或公司清单时，不重复做广泛搜索；
2. 搜索岗位时优先使用结构化 API、公司 career 页面、ATS board，而不是长网页全文读取；
3. 只提取 JD 中与岗位匹配、申请问题、硬性要求相关的内容；
4. Excel 只增量更新，不每次全文重写；
5. 已搜索过的公司只检查新增岗位，不重复整理旧岗位；
6. 能由浏览器插件自动填的基础字段，不消耗 Agent token 重写；
7. 开放题只在用户确认要投递该岗位后生成；
8. 申请状态检查只读取状态相关页面或邮件标题，不做全量分析。

---

## 5. 主 Excel 工作簿需求

### 5.1 工作簿名

在本地路径生成一个 Excel 文件：

```text
C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply\excel\岗位搜索与简历投递.xlsx
```

工作簿名称：

```text
岗位搜索与简历投递
```

### 5.2 Sheet 1：意向投递公司

Sheet 名称：

```text
意向投递公司
```

用途：

- 保存用户想投递的公司；
- 保存 Agent 根据简历和目标行业补充的公司；
- 保存公司 career 页面、LinkedIn/BOSS 链接；
- 作为后续岗位搜索的公司来源表。

建议字段：

| 字段名 | 说明 |
|---|---|
| company_id | 公司唯一 ID，由 Agent 生成，避免重名 |
| 公司名称 | 公司中文名或常用名 |
| Company Name | 英文名 |
| 国家/地区 | 公司/岗位目标地区 |
| 城市 | 分公司或岗位所在城市 |
| 是否目标地区有分公司 | 是/否/待确认 |
| 行业 | 金融科技/咨询/互联网/电商/数据服务等 |
| 岗位大类 | 数据分析/数据科学/机器学习/量化/咨询等 |
| 匹配简历 | 对应简历文件或 profile，例如 profile_en/profile_zh |
| Career 官网链接 | 公司官方 career 页面 |
| ATS 类型 | Greenhouse/Lever/Ashby/Workday/自建/未知 |
| ATS Board 链接 | 如果可识别，填写 ATS board 链接 |
| LinkedIn 直聘链接 | 非中国地区优先填写 |
| BOSS 直聘链接 | 中国境内地区优先填写 |
| 牛客/校招链接 | 中文校招场景可填写 |
| 信息来源 | 用户提供/Agent 搜索/插件发现 |
| 最近检查时间 | 最近一次检查该公司岗位的时间 |
| 备注 | 其他说明 |

规则：

- 如果用户选择中国境内地区，Agent 优先补充：
  - 官网 career 页面；
  - BOSS 直聘相关链接；
  - 牛客/校招链接；
  - 中文招聘入口。
- 如果用户选择非中国地区，Agent 优先补充：
  - 官网 career 页面；
  - LinkedIn 相关岗位或公司 jobs 链接；
  - Greenhouse / Lever / Ashby / Workday 等 ATS 链接。
- 如果用户已经调研过，可直接粘贴公司清单到该 sheet；Agent 只负责补全缺失列和校验链接。

### 5.3 Sheet 2：开放岗位列表

Sheet 名称：

```text
开放岗位列表
```

用途：

- 保存从“意向投递公司”中发现的开放岗位；
- 保存 JD 和投递链接；
- 支持增量更新，只添加新开放岗位或更新状态变化。

建议字段：

| 字段名 | 说明 |
|---|---|
| job_id | 岗位唯一 ID，优先使用 ATS job id，否则由 URL hash 生成 |
| company_id | 对应意向投递公司 sheet 的 company_id |
| 公司名称 | 公司名 |
| 岗位名称 | Job title |
| 岗位大类 | 数据分析/数据科学/ML/量化/咨询等 |
| 国家/地区 | 岗位地区 |
| 城市 | 岗位城市 |
| 工作方式 | onsite/hybrid/remote/unknown |
| 语言 | 中文/英文 |
| JD 摘要 | 精简 JD 摘要 |
| JD 原文链接 | Job description 页面链接 |
| 投递链接 | Application link |
| 来源平台 | 官网/LinkedIn/BOSS/Greenhouse/Lever/Ashby/Workday/其他 |
| 匹配分数 | 0-100 |
| 匹配理由 | 简短说明 |
| 风险/缺口 | 年限不符/签证未知/地点不符/语言要求等 |
| 发现时间 | 首次发现时间 |
| 最近确认开放时间 | 最近一次确认仍开放的时间 |
| 是否已投递 | 是/否/待确认 |
| 用户选择 | 想投/不投/待定 |
| 备注 | 其他说明 |

规则：

- Agent 在用户要求时顺序检查公司是否有符合条件的新开放岗位；
- 如果已检查过一轮，后续只关注新增岗位或已关闭岗位状态变化；
- Agent 应先在对话框中列出候选岗位的 JD 和投递链接，等待用户确认是否投递；
- 未经用户确认，不启动投递填写流程。

### 5.4 Sheet 3：已投递记录

Sheet 名称：

```text
已投递记录
```

用途：

- 记录用户已经实际投递或已完成半自动填写并由用户确认提交的岗位；
- 记录 JD 链接、投递链接、投递时间和状态变化。

建议字段：

| 字段名 | 说明 |
|---|---|
| application_id | 投递记录唯一 ID |
| job_id | 对应开放岗位列表 job_id |
| 公司名称 | 公司名 |
| 岗位名称 | Job title |
| 国家/地区 | 地区 |
| 城市 | 城市 |
| 使用简历 | 使用哪份简历或 profile |
| JD 原文链接 | Job description 链接 |
| 投递链接 | Application link |
| 投递平台 | LinkedIn/BOSS/官网/Greenhouse/Lever/Ashby/Workday 等 |
| 投递时间 | 用户确认已提交的时间 |
| 填表工具组合 | Hermes/autoapply-mcp/牛客/Simplify 等 |
| 是否用户最终提交 | 是/否/待确认 |
| 当前状态 | submitted / under review / assessment / interview / rejected / offer / unknown |
| 状态更新时间 | 最近状态更新时间 |
| 投递状态1 | 第一次状态变化 |
| 更新时间1 | 第一次状态变化时间 |
| 投递状态2 | 第二次状态变化 |
| 更新时间2 | 第二次状态变化时间 |
| 投递状态3 | 第三次状态变化 |
| 更新时间3 | 第三次状态变化时间 |
| 投递状态4 | 第四次状态变化 |
| 更新时间4 | 第四次状态变化时间 |
| 备注 | 用户或 Agent 备注 |

规则：

- 只有当用户明确表示已提交，或用户指示记录为已投递时，才能写入“已投递记录”；
- 如果只是填表到最终提交前，状态应记录为 `filled_not_submitted` 或保留在开放岗位列表中；
- 状态更新采用追加列方式，不覆盖历史状态；
- 第一次检测到状态变化，写入 `投递状态1 / 更新时间1`；第二次写入 `投递状态2 / 更新时间2`，依次类推。

---

## 6. 简历与岗位大类关联需求

### 6.1 用户输入

用户可能提供：

- 一份简历 + 一个岗位大类；
- 多份简历 + 多个岗位大类；
- 一份中文简历 + 一份英文简历；
- 已调研的公司/岗位 Excel；
- 手动粘贴的公司列表。

### 6.2 Agent 处理规则

如果用户提供多份简历并意向投递多个行业与岗位大类，Agent 必须建立：

```text
简历 / profile ↔ 行业 ↔ 岗位大类 ↔ 公司类型 ↔ 地区
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
      "resume_file": "resumes/pdf/Zhou_Tianyi_Resume_EN.pdf",
      "target_regions": ["Singapore", "APAC", "Remote"],
      "industries": ["FinTech", "Consulting", "Internet", "Data Services"],
      "role_categories": ["Data Analyst", "Data Scientist", "Business Analyst"]
    },
    {
      "profile": "profile_zh.json",
      "resume_file": "resumes/pdf/Zhou_Tianyi_Resume_ZH.pdf",
      "target_regions": ["中国大陆"],
      "industries": ["金融科技", "咨询", "电商", "互联网"],
      "role_categories": ["数据分析", "商业分析", "技术咨询", "投研数据"]
    }
  ]
}
```

### 6.3 匹配原则

Agent 推荐公司和岗位时，应综合：

- 专业背景匹配度；
- 技能匹配度；
- 实习经历匹配度；
- 项目/竞赛经历匹配度；
- 地理位置匹配度；
- 是否有目标国家/地区分公司；
- 语言匹配度；
- 年限要求匹配度；
- 签证/工作授权风险；
- 用户明确排除的行业或公司。

---

## 7. 岗位搜索流程需求

### 7.1 可选入口 A：Agent 主动搜索

当用户只提供简历和目标行业/岗位大类时，Agent 应：

1. 向用户确认目标国家/地区；
2. 解析简历并建立 profile；
3. 按地区和岗位大类搜索合适公司；
4. 检查公司是否在目标地区有分公司或开放岗位；
5. 补充 career 页面和 LinkedIn/BOSS 链接；
6. 写入 Excel 的“意向投递公司”sheet；
7. 在用户要求时继续搜索开放岗位并写入“开放岗位列表”。

### 7.2 可选入口 B：用户上传或粘贴公司/岗位 Excel

当用户已经调研过并上传公司或岗位列表时，Agent 应：

1. 读取用户提供的 Excel；
2. 判断它更接近“意向投递公司”还是“开放岗位列表”；
3. 映射到主工作簿 schema；
4. 只补全缺失信息，不随意覆盖用户原始内容；
5. 如需修改用户已有内容，先说明差异并请求确认。

### 7.3 岗位增量搜索

当用户要求查找岗位，或定时任务触发时：

1. 顺序读取“意向投递公司”；
2. 对每家公司检查 career 页面、ATS board、LinkedIn/BOSS 入口；
3. 找出符合用户行业与岗位大类的新开放岗位；
4. 与“开放岗位列表”中已有 `job_id` 或 URL 去重；
5. 只新增新岗位，或更新已关闭/状态变化字段；
6. 在对话框中汇报新增岗位 JD 摘要和投递链接；
7. 等用户确认后才启动申请填写。

---

## 8. 半自动投递流程需求

### 8.1 启动条件

半自动投递流程只能在以下条件满足时启动：

1. 用户明确选择要投递的岗位；
2. 已确定使用哪份简历/profile；
3. 已确认目标语言：中文或英文；
4. 已确认投递平台和工具组合；
5. 敏感字段缺口已列出；
6. 用户知道 Agent 不会点击最终提交。

### 8.2 工具组合选择

#### 中文网站

默认：

```text
Hermes browser automation + 牛客网申助手
```

备用：

```text
Hermes browser automation only
```

可选：

```text
OfferLink 等中文插件
```

#### 英文网站 / LinkedIn

默认：

```text
Hermes browser automation + Simplify Copilot
```

ATS 专项：

```text
autoapply-mcp + Hermes browser automation
```

备用：

```text
Hermes browser automation only
```

### 8.3 自动填写范围

允许自动填写：

- 姓名；
- 邮箱；
- 电话，但若有多个号码需用户确认；
- 教育经历；
- 实习经历；
- 项目经历；
- 技能；
- 语言能力；
- 简历上传；
- 基于真实经历生成的开放题草稿。

不允许自动猜测：

- 身份证号；
- 政治面貌；
- 户籍；
- 期望薪资；
- 到岗时间；
- 签证状态；
- 工作授权；
- 是否需要 sponsorship；
- 是否接受调剂；
- 是否有亲属任职；
- EEO / disability / veteran / demographic 信息；
- 背景调查、犯罪记录、法律声明；
- 真实性承诺。

### 8.4 提交边界

硬性规则：

```text
Agent 永远不自动点击最终 Submit / Apply / Submit Application / 提交申请按钮。
```

Agent 必须在最终提交前：

1. 汇报已填写字段；
2. 汇报未填写或需确认字段；
3. 汇报开放题草稿来源；
4. 提醒用户自行检查；
5. 等待用户自行提交或放弃。

---

## 9. 定时任务需求

### 9.1 每日岗位搜索 Timer

用户希望设置 timer：

```text
当地时间每天早晨 10:00
```

触发任务：

1. 读取“意向投递公司”sheet；
2. 检查每家公司是否有新开放岗位；
3. 只关注用户目标行业和岗位大类；
4. 如果已有历史搜索记录，只汇报新开放岗位；
5. 将新岗位写入“开放岗位列表”；
6. 在对话框中给出新岗位 JD 摘要和投递链接；
7. 若用户接受某岗位，启动半自动投递流程。

### 9.2 每日投递状态检查 Timer

同一 timer 或另一个 timer 在每天 10:00 触发：

1. 读取“已投递记录”sheet；
2. 检查已投递岗位是否有状态更新；
3. 可能状态来源包括：
   - 申请门户；
   - ATS 页面；
   - 用户授权后的邮件状态通知；
   - 手动提供的状态更新；
4. 如果有状态变化，则写入新的 `投递状态N / 更新时间N`；
5. 如果没有状态变化，不改动 Excel；
6. 如果需要登录或验证码，暂停并请求用户介入。

### 9.3 Timer 安全规则

- Timer 不自动投递；
- Timer 不自动登录；
- Timer 不自动填写敏感字段；
- Timer 不删除 Excel 中任何数据；
- Timer 的输出必须可继续交互，让用户确认是否投递；
- Timer prompt 必须是自包含的，不依赖当前聊天上下文。

---

## 10. Excel 数据修改与权限规则

### 10.1 禁止私自删除

未经用户明确允许，Agent 不得删除 Excel 内任何信息。

包括：

- 不删除公司；
- 不删除岗位；
- 不删除历史状态；
- 不删除用户备注；
- 不删除旧链接；
- 不删除手动粘贴内容。

### 10.2 更新规则

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

### 10.3 审计记录

每次写入 Excel 时，应在日志中记录：

- 时间；
- 操作类型；
- 修改 sheet；
- 新增/更新的行数；
- 数据来源；
- 是否经用户确认；
- 相关岗位链接。

---

## 11. 用户交互需求

Agent 需要在关键节点主动询问用户，而不是猜测。

### 11.1 初始确认

- 投递国家/地区；
- 中文流程还是英文流程；
- 使用哪份简历；
- 目标行业；
- 目标岗位大类；
- 是否有不投公司/行业；
- 是否已有公司/岗位 Excel。

### 11.2 投递前确认

- 选择哪些岗位投递；
- 使用哪份简历/profile；
- 是否使用插件辅助；
- 电话号码选择；
- 当前所在地；
- 工作授权/签证；
- 薪资/到岗时间等敏感字段；
- 开放题草稿是否可用。

### 11.3 投递后确认

- 用户是否已经手动提交；
- 是否记录为已投递；
- 是否需要记录备注；
- 是否需要安排后续状态检查。

---

## 12. 安全与合规要求

硬性要求：

1. 不自动点击最终提交；
2. 不处理账号密码；
3. 不破解 CAPTCHA；
4. 不使用代理池、指纹伪装或反检测工具；
5. 不调用 LinkedIn 非官方 API 作为主方案；
6. 不伪造经历、技能、教育背景；
7. 不自动回答敏感字段；
8. 不删除 Excel 历史信息；
9. 不未经确认覆盖用户手动输入；
10. 所有申请内容必须基于简历或用户确认材料。

---

## 13. MVP 范围

### 13.1 MVP 必须实现

1. 在项目目录下生成主 Excel：`岗位搜索与简历投递.xlsx`；
2. 建立至少三个 sheet：
   - `意向投递公司`；
   - `开放岗位列表`；
   - `已投递记录`；
3. 从用户简历生成或维护：
   - `profile_en.json`；
   - `profile_zh.json`；
4. 支持用户手动提供公司/岗位 Excel 并映射进主工作簿；
5. 支持按公司查找开放岗位；
6. 支持用户确认岗位后启动半自动填表；
7. 支持填写前敏感字段清单；
8. 支持在最终提交前停止；
9. 支持记录已投递岗位；
10. 支持每日 10:00 的岗位更新和状态检查 timer 设计。

### 13.2 MVP 暂不实现

1. 不实现全自动提交；
2. 不实现 LinkedIn 非官方 API；
3. 不实现验证码破解；
4. 不实现自动批量海投；
5. 不强制接入第三方 CRM；
6. 不立即安装 autoapply-mcp 或浏览器插件，先完成需求与设计。

---

## 14. 验收标准

### 14.1 文件与路径

- 所有项目文件均在 `C:\Users\24779\Desktop\AI related\Case A_Job Auto Apply` 下；
- Excel、profiles、logs、applications、mcp 目录结构清晰；
- 不将项目数据散落到无关目录。

### 14.2 Excel 工作簿

- 存在 `岗位搜索与简历投递.xlsx`；
- 至少包含 `意向投递公司`、`开放岗位列表`、`已投递记录`；
- 字段满足岗位搜索和投递记录需求；
- 新增和状态更新可追溯；
- 不经用户允许不会删除历史数据。

### 14.3 工作流

- Agent 能根据用户简历和目标岗位大类整理公司；
- Agent 能检查新开放岗位；
- Agent 能给出 JD 摘要和投递链接；
- Agent 能在用户确认后顺序启动半自动投递；
- Agent 能根据中文/英文选择不同 profile 和插件辅助；
- Agent 能在最终提交前停止；
- Agent 能记录投递结果。

### 14.4 安全

- 敏感字段不会被猜测；
- 登录和密码由用户自己处理；
- 不自动提交；
- 不绕过平台限制；
- 申请内容不编造。

---

本需求文档仅用于后续开发整个半自动简历投递工作流的设计依据。
