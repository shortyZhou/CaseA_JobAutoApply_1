# Job Auto Apply（岗位搜索与简历投递自动化）

一个半自动化求职投递工作流项目：由 Agent 结合 **Hermes browser automation** 与 **autoapply-mcp** 工具，辅助完成公司调研、岗位搜索、半自动投递与投递状态追踪。中文网站投递配合「牛客网申助手」浏览器插件，英文网站投递配合「Simplify Copilot」浏览器插件。

## 核心功能（Functional）

1. **生成投递总表** — 在本地生成 Excel《岗位搜索与简历投递》，先与用户确认投递地区（中国境内 / 非中国地区）。
2. **简历-公司匹配调研** — 用户提交简历并描述意向行业与岗位大类；Agent 为每对「简历 → 公司/岗位大类」建立关联，找出地理上符合投递地区的目标公司，并填入《岗位搜索与简历投递》的 **「意向投递公司」** sheet：
   - 中国境内：公司官网 career 页面链接 + BOSS 直聘链接
   - 非中国地区：公司官网 career 页面链接 + LinkedIn 链接
   - 用户已有目标公司时，可直接粘贴进该 sheet，由 Agent 补全信息。
3. **岗位搜索** — 按需顺序检查「意向投递公司」中各公司是否有符合意向的开放/新开放岗位，给出 Job Description 与投递链接；已查过的公司只关注新开放岗位。
4. **半自动投递** — 用户确认要投的岗位后，Agent 按顺序启动半自动投递流程，并记录到 **「已投递记录」** sheet（公司、岗位、JD 链接、投递时间）。
5. **每日定时任务（当地时间早 10 点）**
   - 搜索新开放岗位，给出 JD 与投递链接；用户接受后进入半自动投递流程并记录。
   - 检查已投递岗位的状态更新；有更新则在「已投递记录」追加 `投递状态N`、`更新时间N` 列（N 按第几次检测顺次编号），无更新则不改动。
6. **数据安全规则** — 未经用户允许，不得删除或更新 Excel 中的任何已有信息。

## 重要路径说明

### 简历放置位置（Resume）

| 位置 | 用途 |
|---|---|
| `resumes/original/` | 用户提供的原始简历 |
| `resumes/pdf/` | 转换后的 PDF 版简历 |
| `.autoapply/resumes/` | **autoapply-mcp 投递流程实际读取的简历**，路径在 `.autoapply/profile.json` 的 `profile.resumes[].path` 中配置（需为绝对路径，推荐正斜杠写法，PDF ≤ 25MB） |

> 新增或更换简历时：先放入 `resumes/`，再更新 `.autoapply/profile.json` 中的 `profile.resumes[].path`（或直接复制到 `.autoapply/resumes/` 并同步路径），确保投递时能校验通过（文件存在、非空、有效 PDF 签名）。

### Excel 生成位置

- 工作簿：`excel/岗位搜索与简历投递.xlsx`
- 三个核心 sheet：
  - **意向投递公司** — 目标公司及 career / BOSS 直聘 / LinkedIn 链接
  - **开放岗位列表** — 搜索到的岗位明细
  - **已投递记录** — 投递记录及状态更新（`投递状态N` / `更新时间N`）
- 每次写入前自动生成时间戳备份至 `backups/excel/`，已有用户填写内容不会被覆盖。

## 目录结构

```
├── resumes/                # 简历（original 原始件 / pdf 转换件）
├── excel/                  # 《岗位搜索与简历投递.xlsx》输出位置
├── profiles/               # 用户资料、投递偏好、简历-岗位映射
├── applications/           # 申请 packet 数据
├── scripts/                # Excel 维护、简历转换、状态检查等脚本
├── logs/                   # 运行日志（job search / autofill / status check）
├── backups/                # Excel 等文件的时间戳备份
├── .autoapply/             # autoapply-mcp 运行数据（profile.json、公司列表、数据库、投递用简历）
└── mcp/autoapply-mcp/      # autoapply-mcp 源码（TypeScript MCP server）
```

## Excel 维护命令示例

```bash
python scripts/update_job_excel.py summary
python scripts/update_job_excel.py append-job --json '{"job_id":"demo","Company Name":"DemoCo","Job Title":"Analyst"}'
python scripts/update_job_excel.py mark-applied --job-id demo --submitted Yes --status submitted
python scripts/update_job_excel.py update-status --job-id demo --status "under review"
```
