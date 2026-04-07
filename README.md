# dcd-koubei-collector

懂车帝车型口碑采集 skill。

基于懂车帝车型口碑分页页面抓取用户评价，导出 Excel 与校验报告，适合做车型口碑整理、竞品分析、人工复核和后续摘要分析。

## 当前能力

- 抓取懂车帝车型口碑分页数据
- 自动探测总页数
- 导出 `xlsx`
- 同步生成 `.validation.json`
- 同步生成 `.progress.json`
- 同步生成 `.failed-pages.json`
- 终端文本进度条，展示阶段 / 页码 / 成功重试失败 / 累计记录数
- 支持把进度实时写入 JSON、推送到通用 webhook，或直连飞书 incoming webhook
- 支持按失败页补抓：`--retry-failed-pages`
- 支持把补抓结果合并进旧结果：`--merge-into`
- 支持合并模式：`--merge-mode keep-extra|strict`
- 当前支持字段：
  - 用户名
  - 用户标签
  - 评价车型
  - 懂车分
  - 发布时间
  - 用户评分
  - 续航
  - 购车时间
  - 裸车价
  - 购车地
  - 评价全文
  - 来源链接
  - 抓取页码

## 目录结构

```text
.
├── SKILL.md
├── README.md
├── references/
│   └── rules.md
└── scripts/
    └── export_dcd_koubei.py
```

## 用法

### 方式 1：传 `series-id`

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --start-page 1
```

### 方式 2：传懂车帝车型口碑页 URL

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --url 'https://www.dongchedi.com/auto/series/score/25544-x-x-x-x-x' \
  --start-page 1
```

### 指定输出文件

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --start-page 1 \
  --end-page 3 \
  --output ./DCD口碑_风云X3_2026-03-27.xlsx
```

## 默认输出

如果不传 `--output`，默认输出：

```text
DCD口碑_车型_日期.xlsx
```

例如：

```text
DCD口碑_风云X3_2026-03-27.xlsx
```

同时会生成这些同名侧产物：

```text
DCD口碑_风云X3_2026-03-27.validation.json
DCD口碑_风云X3_2026-03-27.progress.json
DCD口碑_风云X3_2026-03-27.failed-pages.json
```

## 数据来源说明

当前实现直接解析页面 HTML 中的 `__NEXT_DATA__` 结构化数据，不依赖 OCR。

这意味着：

- 速度比浏览器逐条抠 DOM 更稳
- 字段结构更规整
- 但如果懂车帝未来改页面注水方式，这个解析逻辑也需要一起更新

## 仓库维护方式

这个仓库作为独立 skill 仓库维护，并在 OpenClaw workspace 中以 submodule 方式接入。

- 独立仓库工作目录：`repos/dcd-koubei-collector`
- workspace submodule 路径：`skills/dcd-koubei-collector`

日常更新建议：

1. 在独立仓库目录修改并提交
2. push 到 GitHub
3. 回 workspace 根目录提交 submodule 指针更新

## 增强用法

### 进度展示

运行时会输出类似：

```text
抓取页面 [########................] 总体 8/21 (38%) | 页码 8/20 | ok 7 retry 0 fail 0 rows 134 | 第 8 页
```

如果要把进度写给对话框或前端轮询，可加：

- `--progress-file /tmp/dcd.progress.json`

如果要主动推送到通用 webhook，可加：

- `--progress-webhook https://...`

如果要直连飞书 incoming webhook，可加：

- `--feishu-webhook https://...`
- `--feishu-secret <secret>`

### 失败页补抓

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json
```

### 合并进旧结果

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json \
  --merge-into ./DCD口碑_旧版本.xlsx \
  --output ./DCD口碑_修复版.xlsx
```

### 严格合并模式

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json \
  --merge-into ./DCD口碑_旧版本.xlsx \
  --merge-mode strict \
  --output ./DCD口碑_修复版.xlsx
```

- `keep-extra`：默认，保留旧表中本轮未触及的记录
- `strict`：只保留本轮新结果，不保留历史残留

## 当前限制

- 目前未接入更深层详情页补字段逻辑
- `用户评分` 当前与 `懂车分` 共用同一结构化分值
- 如页面字段缺失，导出时留空，不强行补值

## 后续可扩展方向

- 点赞数
- 评论数
- 图片数
- 用户 ID
- 详情页补充字段
- 更细的车型/版本筛选


## Release 内容

每个 Release 默认包含：

- `.skill` 包
- GitHub 自动生成的源码压缩包

## 仓库

- GitHub: <https://github.com/sh3rlockC/dcd-koubei-collector>
