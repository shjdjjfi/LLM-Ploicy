# 数据集构造脚本（含 19 个 `delta_*` 列 + `gov_expected_changes`）

由于当前执行环境无法直接访问你提供的 Google Drive 链接（网络代理限制），我在仓库中提供了可直接落地的自动化脚本。你把本地文件路径替换后即可一键生成结果。

## 目录

- `scripts/build_dataset.py`：主脚本
- `mapping_template.csv`：19 个 `delta_*` 列到世界银行指标代码的映射模板

## 你需要准备

1. 世界银行指标 zip（1960-2023）
2. 样本指引 xlsx（需要新增 20 列）
3. 填写好的 `mapping.csv`（可基于 `mapping_template.csv`）

`mapping.csv` 格式：

```csv
delta_col,indicator_code,direction
```

- `delta_col`：xlsx 中目标列名（例如 `delta_gdp_growth`）
- `indicator_code`：世界银行指标代码（例如 `NY.GDP.MKTP.KD.ZG`）
- `direction`：可选，默认 `1`；若希望“下降是改善”，可填 `-1`

## 运行方式

```bash
python scripts/build_dataset.py \
  --wb-zip /path/to/world_bank.zip \
  --guide-xlsx /path/to/sample_guide.xlsx \
  --mapping-csv /path/to/mapping.csv \
  --output /path/to/output_with_deltas.xlsx
```

可选参数：

- `--sheet`：工作表名（默认第一个）
- `--country-col`：国家列名（默认自动匹配）
- `--year-col`：年份列名（默认自动匹配）
- `--agg`：`mean` 或 `sum`，用于聚合 19 个 delta 到 `gov_expected_changes`（默认 `mean`）

## 计算逻辑

- `delta_* = direction * (value(year) - value(year-1))`
- `gov_expected_changes = agg(所有可用 delta_* 值)`

当某一项缺失时，该 `delta_*` 留空；`gov_expected_changes` 会基于可用项计算。
