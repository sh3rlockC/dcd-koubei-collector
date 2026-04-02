---
name: dcd-koubei-collector
description: 从懂车帝口碑页批量采集指定车型的用户评价全文，并导出为 Excel。适用于用户要求抓取懂车帝口碑、整理车型评价、按分页批量采集、输出表格、保留评价原文、提取评分/价格/购车信息等场景。尤其在用户给出懂车帝口碑链接、车型页、分页范围，或明确提到“懂车帝”“口碑”“评价全文”“导出 Excel/表格”时使用。
---

# 懂车帝车型口碑采集

按下面流程执行，目标是稳定抓取懂车帝车型口碑，并输出可复核、可分析的 Excel。

## 1. 确认输入

优先接受三种输入：

- 完整懂车帝口碑页 URL
- `车型页链接 + 页码范围`
- 自然语言任务描述（如“抓懂车帝某车型口碑 1-20 页导出 Excel”）

页码处理规则：

- 如果用户明确给了页码范围，就按用户指定页码抓取
- 如果用户没有给页码，或明确说“全部页面 / 所有页面 / 所有条目”，则先检测分页总页数，再按 `1 ~ 最后一页` 抓取
- 若用户给了 `start-page` 但没给 `end-page`，脚本默认可自动探测 `end-page`

## 2. 抓取方式

优先使用浏览器方式读取页面可见文本，不要默认走 OCR。

执行原则：

- 先打开列表页
- 从列表页提取结构字段
- 评价全文优先取列表页展示文本
- 若列表页某条记录缺正文，则自动回退详情页补全文
- 正式开发或高精度任务前，抽样对比详情页正文是否更完整
- 若详情页整体更完整，则改为“列表页取结构字段 + 详情页补全文”

## 3. 正式版字段

输出以下字段：

1. 评价车型
2. 懂车分
3. 发布时间
4. 用户评分
5. 续航
6. 购车时间
7. 裸车价
8. 购车地
9. 评价全文
10. 来源链接
11. 抓取页码

字段原则：

- `评价全文` 必须保留网页原文，不删改、不摘要、不纠错
- 页面无字段时留空，不强行补值
- `裸车价` 保留页面原始展示内容
- `购车地` 保留页面原始展示内容

## 4. 与汽车之家版的关键差异

不要沿用“最满意 / 最不满意”双维度逻辑。

本 skill 默认：

- 不区分“最满意 / 最不满意”
- 直接按单条口碑抓取 `评价全文`
- 不做双维度链接对齐
- 默认输出单一明细表

## 5. 导出规则

默认文件命名：

- `DCD口碑+车型+日期.xlsx`
- 例如：`DCD口碑_小米SU7_2026-03-27.xlsx`
- 若用户显式要求保留页数范围，可额外追加页数后缀

默认输出 1 个 sheet：

- `口碑明细`

若用户后续要求，也可扩展为：

- 原始明细
- 清洗明细
- 异常记录

## 6. 校验规则

每条记录至少应包含以下核心字段：

- 评价车型
- 发布时间
- 评价全文
- 来源链接

缺任意一项时：

- 标记为异常记录
- 不直接写入正式结果
- 记录到异常清单或 JSON 侧产物中

每页抓取后检查：

- 是否成功拿到口碑列表区
- 是否识别到合理数量的口碑卡片
- 是否出现大面积字段为空

正式导出后，必须生成同名 `.validation.json` 报告，至少包含：

- 总抓取条数
- 各分页条数
- 异常条数
- 异常原因摘要

## 7. 重试规则

分页抓取失败时：

- 自动重试 2 次
- 连续失败后停止静默继续，明确报告页码和错误原因

## 8. 实操提醒

- 先确认懂车帝页面是否存在登录、滑块、接口签名或反爬限制
- 不要因为文案口语化、像广告、像模板文就擅自删除正文
- 批量任务除 Excel 外，最好同步保留 JSON 原始结构化结果，方便回溯和修规则
- 正式开发前，至少抽样对比 5 条列表页正文与详情页正文

## 9. 当前阶段目标

当前版本已完成可用版落地。

当前目标：

- 稳定抓取懂车帝车型口碑分页数据
- 导出结构化 Excel 与 validation.json
- 在后续迭代中继续补充用户名、车主标签等增强字段

## 10. 参考资料

如需快速查看稳定规则，不必重复翻全文，直接看：

- `references/rules.md`

## 11. 附带脚本

本 skill 附带脚本：

- `scripts/export_dcd_koubei.py`

当前用途：

- 直接抓取懂车帝车型口碑分页数据并导出 Excel
- 支持从页面内嵌结构化数据提取字段
- 支持自动探测总页数
- 默认文件名为 `DCD口碑_车型_日期.xlsx`
- 运行时默认输出文本进度条，展示阶段、页码进度、成功/重试/失败页数、累计记录数
- 默认额外生成同名 `.progress.json`，便于外部读取实时进度
- 默认额外生成同名 `.failed-pages.json`，集中记录失败页和失败原因，方便补抓
- 支持通过 `--retry-failed-pages` 直接读取失败页并补抓
- 支持通过 `--merge-into` 将补抓结果按 `来源链接` 合并进已有 Excel，覆盖旧记录
- 支持通过 `--merge-mode keep-extra|strict` 控制是否保留旧表中本轮未触及的历史记录
- 推荐合并后输出到新文件（如 `_修复版.xlsx`），避免直接覆盖原始产物
- 可通过 `--progress-file` 自定义进度文件路径
- 可通过 `--quiet` 关闭终端进度输出，仅保留 `.progress.json`

示例：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --start-page 1
```

运行时会输出类似：

```text
抓取页面 [########................] 总体 8/21 (38%) | 页码 8/20 | ok 7 retry 0 fail 0 rows 134 | 第 8 页
```

或：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --url 'https://www.dongchedi.com/auto/series/score/25544-x-x-x-x-x' \
  --start-page 1
```

`.progress.json` 会包含：

- `overall.current / total / percent`
- `current_page`
- `page_range`
- `output_path`
- `validation_path`
- `failed_pages`

如需按失败页补抓：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json
```

如需把补抓结果合并进已有 Excel：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json \
  --merge-into ./DCD口碑_小米SU7_2026-03-27.xlsx \
  --output ./DCD口碑_小米SU7_2026-03-27_修复版.xlsx
```

如需严格模式（不保留旧表中本轮未触及的历史记录）：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --retry-failed-pages ./DCD口碑_xxx.failed-pages.json \
  --merge-into ./DCD口碑_小米SU7_2026-03-27.xlsx \
  --merge-mode strict \
  --output ./DCD口碑_小米SU7_2026-03-27_修复版.xlsx
```

如需静默运行，只保留 progress 文件：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --start-page 1 \
  --quiet
```

如需自定义 progress 文件：

```bash
python3 skills/dcd-koubei-collector/scripts/export_dcd_koubei.py \
  --series-id 25544 \
  --start-page 1 \
  --progress-file ./dcd.progress.json
```
