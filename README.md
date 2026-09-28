# wide-table-etl

多来源 Excel 横表自动转标准长表（Python 数据自动化项目 · 作品集演示版）

多个业务单元每天上报格式各异的 Excel 横表（列数不同、日期表示不同、表头位置不同），
本工具按模板规则自动定位数据区域、宽转长、统一口径后汇总出标准数据集。
项目流程抽象自一个运行多年的多基地库存报表汇总场景，**仓库内全部数据均为程序合成，不含任何真实业务信息**。

姊妹项目：https://github.com/GenGC0128/coa-portfolio（异构 PDF 批量解析），同一套“规则驱动解析”架构。

**Lead: @GenGC0128** 

## 特性

- **模板规则配置化**：每类横表一条 `BaseRule`（日期行/identify 行/供应商列/规格列/日期模式/表尾标志），新增来源只加配置
- **动态边界定位**：日期起点自动识别、右边界支持关键字定位或行尾自适应、表尾按标志词/空行截断——报表加列加行不影响解析
- **三种日期模式统一**：datetime / Excel serial（`origin='1899-12-30'`）/ 行内日期
- **口径归一化外置**：规格/密度/粒度/类别/供应商词典在 `config/normalization.json`，业务规则与代码分离
- **失败隔离**：单文件异常不中断批次，失败原因落盘 `failed.csv`

## 快速开始

```bash
pip install -r requirements.txt
python make_sample_data.py
python inventory_pipeline.py --input data/reports --config config/normalization.json --out data/output
pytest tests/ -v
```

输出：`data/output/summary.xlsx`、`report.csv`、`failed.csv`

## 架构

```
data/reports/<source>/*.xlsx
        │
        ▼
extract_wide()        # 按 BaseRule 动态定位边界，横表转长表（统一 schema）
        ▜─ datetime / serial / 行内日期 三种模式
        ▜─ project_mode 支持无 identify 的项目类报表
        ▼
normalize()           # config/normalization.json 词典：spec/density/size/category/supplier
        ▼
summary.xlsx + report.csv + failed.csv
```

## 设计决策备忘

| 问题 | 方案 |
|---|---|
| 各来源表格坐标各异 | 坐标即配置：模板差异收敛到 `BaseRule` |
| 报表列数/行数经常变化 | 动态定位边界，不写死行列号 |
| Excel 序列日期易错 | 统一 `pd.to_datetime(unit='D', origin='1899-12-30')`，不手搓 timedelta |
| 口径散落在代码 if-else | 归一化词典外置 JSON，改规则不动代码 |
| 历史脚本 notebook 化难维护 | CLI 参数化 + 纯函数阶段拆分 + pytest 端到端测试 |

## License

MIT（见 LICENSE）
