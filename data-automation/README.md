# 🤖 Excel 数据自动化处理工具

> Python 自动化演示项目 | 适用于 Fiverr 展示

## 功能亮点

- ✅ **批量处理**：自动扫描文件夹，合并所有 Excel/CSV
- ✅ **智能汇总**：自动识别数值列，计算总和/均值/最大/最小值
- ✅ **多 Sheet 导出**：合并数据 + 按文件汇总 + 错误日志
- ✅ **演示数据**：首次运行自动生成示例文件，客户可直接体验
- ✅ **错误容错**：单个文件损坏不影响整体处理

## 适用场景（Fiverr 接单方向）

| 行业 | 具体需求 |
|------|---------|
| 电商 | 多店铺销售报表合并 + 利润分析 |
| 财务 | 各部门预算汇总、发票数据清洗 |
| 人事 | 考勤统计、工资计算 |
| 教育 | 学生成绩合并、排名统计 |
| 物流 | 发货记录汇总、运费核算 |

## 技术栈

| 技术 | 用途 |
|------|------|
| `pandas` | 数据读取、清洗、分析 |
| `openpyxl` | Excel 多 Sheet 写入 |
| `pathlib` | 跨平台文件路径处理 |

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行（自动生成演示数据 + 处理 + 导出报告）
python excel_processor.py

# 3. 查看结果
ls output/
# → summary_report.xlsx  汇总报告（多 Sheet）
# → detailed_data.csv    合并后的详细数据
```

## 输出示例

```
🚀 Excel 自动化处理工具启动
📝 生成演示数据...
✅ 演示数据已生成到 input/ 文件夹
📂 找到 2 个文件待处理
  ✓ sales_january.xlsx → 5 行
  ✓ sales_february.xlsx → 5 行
✅ 合并完成: 10 行, 7 列
==================================================
📊 汇总报告
   总记录: 10
   文件数: 2
   [quantity] 总和=1265.0, 均值=126.5
   [unit_price] 总和=455.0, 均值=45.5
   [total] 总和=38207.5, 均值=3820.75
==================================================
📄 详细数据: output/detailed_data.csv
📊 汇总报告: output/summary_report.xlsx
🎉 处理完成！
```

## 目录结构

```
data-automation/
├── excel_processor.py   # 主程序
├── requirements.txt     # 依赖
├── input/              # 放待处理的 Excel（会自动创建演示数据）
│   ├── sales_january.xlsx
│   └── sales_february.xlsx
└── output/             # 输出结果
    ├── summary_report.xlsx
    └── detailed_data.csv
```

## 改造指南（适配不同客户需求）

```python
# 场景 1：只处理 CSV
excel_files = list(self.input_dir.glob("*.csv"))  # 删掉 xlsx 那行

# 场景 2：自定义统计维度
def generate_report(self, df):
    df['月'] = pd.to_datetime(df['date']).dt.month
    monthly = df.groupby('月')['total'].sum()   # 按月汇总
    ...

# 场景 3：添加数据校验
def _validate(self, df):
    if df['quantity'].min() < 0:
        raise ValueError("数量不能为负数！")
```

---

👨‍💻 **Fiverr 接单时**：把输入文件夹截图 + output summary 截图发给客户，效果拉满。
