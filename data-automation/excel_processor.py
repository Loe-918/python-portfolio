"""
Excel 数据自动化处理工具
==========================
演示用途：批量处理销售报表 Excel 文件，自动生成汇总报告。
实际接单时可改造成：财务报表合并、库存统计、员工考勤汇总 等。

作者：Liam (Fiverr Portfolio)
"""

import glob
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

# ============== 配置 ==============
INPUT_DIR = "input"                # 输入文件夹（放待处理的 Excel）
OUTPUT_DIR = "output"              # 输出文件夹
SUMMARY_FILE = "summary_report.xlsx"
DETAIL_FILE = "detailed_data.csv"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


class ExcelProcessor:
    """专业的 Excel 批量处理器"""

    def __init__(self, input_dir: str = INPUT_DIR, output_dir: str = OUTPUT_DIR):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.all_data: list[pd.DataFrame] = []
        self.errors: list[dict] = []

    # ---------- 核心处理 ----------
    def process_all(self) -> pd.DataFrame:
        """处理 input/ 下所有 Excel/CSV 文件"""
        excel_files = list(self.input_dir.glob("*.xlsx")) + list(self.input_dir.glob("*.xls"))
        csv_files = list(self.input_dir.glob("*.csv"))
        all_files = excel_files + csv_files

        if not all_files:
            logger.warning(f"⚠️ {self.input_dir}/ 下没有找到 Excel/CSV 文件")
            return pd.DataFrame()

        logger.info(f"📂 找到 {len(all_files)} 个文件待处理")

        for file_path in all_files:
            self._process_file(file_path)

        if not self.all_data:
            logger.error("❌ 没有成功处理任何文件")
            return pd.DataFrame()

        merged = pd.concat(self.all_data, ignore_index=True)
        logger.info(f"✅ 合并完成: {len(merged)} 行, {len(merged.columns)} 列")
        return merged

    def _process_file(self, file_path: Path):
        """处理单个文件"""
        try:
            if file_path.suffix == ".csv":
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # 标准化列名（去空格、统一小写）
            df.columns = df.columns.str.strip().str.lower()

            # 添加数据来源标记
            df["source_file"] = file_path.name
            df["processed_at"] = datetime.now().isoformat()

            self.all_data.append(df)
            logger.info(f"  ✓ {file_path.name} → {len(df)} 行")

        except Exception as e:
            logger.error(f"  ✗ {file_path.name} 处理失败: {e}")
            self.errors.append({
                "file": str(file_path),
                "error": str(e),
                "time": datetime.now().isoformat(),
            })

    # ---------- 报告生成 ----------
    def generate_report(self, df: pd.DataFrame) -> dict:
        """生成汇总报告"""
        if df.empty:
            return {}

        # 自动检测数值列
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        # 排除 source_file 这类标记列（如果有的话）
        numeric_cols = [c for c in numeric_cols if c not in ("source_file",)]

        report = {
            "总记录数": len(df),
            "文件数": df["source_file"].nunique() if "source_file" in df.columns else 1,
            "数值列统计": {},
            "缺失值": df.isnull().sum().to_dict(),
        }

        for col in numeric_cols:
            report["数值列统计"][col] = {
                "总和": round(df[col].sum(), 2),
                "平均值": round(df[col].mean(), 2),
                "最大值": round(df[col].max(), 2),
                "最小值": round(df[col].min(), 2),
            }

        return report

    # ---------- 导出 ----------
    def save_results(self, df: pd.DataFrame):
        """导出处理结果"""
        if df.empty:
            return

        # 详细数据 CSV
        csv_path = self.output_dir / DETAIL_FILE
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        logger.info(f"📄 详细数据: {csv_path}")

        # 汇总报告 Excel（多 Sheet）
        xlsx_path = self.output_dir / SUMMARY_FILE
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            # Sheet 1: 合并后的全部数据
            df.to_excel(writer, sheet_name="合并数据", index=False)

            # Sheet 2: 按来源文件的统计
            if "source_file" in df.columns:
                df.groupby("source_file").agg({
                    col: ["count", "sum", "mean"] for col in df.select_dtypes("number").columns
                    if col != "source_file"
                }).to_excel(writer, sheet_name="按文件汇总")

            # Sheet 3: 错误日志（如果有）
            if self.errors:
                pd.DataFrame(self.errors).to_excel(writer, sheet_name="错误日志", index=False)

        logger.info(f"📊 汇总报告: {xlsx_path}")


# ============== 示例：生成演示数据 ==============
def create_demo_data():
    """生成演示 Excel 文件，方便项目展示"""
    demo_dir = Path(INPUT_DIR)
    demo_dir.mkdir(parents=True, exist_ok=True)

    logger.info("📝 生成演示数据...")

    # 销售数据 1月
    pd.DataFrame({
        "product": ["Widget A", "Widget B", "Widget C", "Widget D", "Widget E"],
        "category": ["Electronics", "Home", "Electronics", "Office", "Home"],
        "quantity": [120, 85, 200, 60, 150],
        "unit_price": [25.00, 45.50, 15.00, 120.00, 32.00],
        "total": [3000, 3867.5, 3000, 7200, 4800],
        "date": ["2024-01-15"] * 5,
    }).to_excel(demo_dir / "sales_january.xlsx", index=False)

    # 销售数据 2月
    pd.DataFrame({
        "product": ["Widget A", "Widget B", "Widget C", "Widget D", "Widget F"],
        "category": ["Electronics", "Home", "Electronics", "Office", "Electronics"],
        "quantity": [95, 110, 180, 45, 220],
        "unit_price": [25.00, 45.50, 15.00, 120.00, 18.00],
        "total": [2375, 5005, 2700, 5400, 3960],
        "date": ["2024-02-15"] * 5,
    }).to_excel(demo_dir / "sales_february.xlsx", index=False)

    logger.info("✅ 演示数据已生成到 input/ 文件夹")


# ============== 入口 ==============
def main():
    logger.info("🚀 Excel 自动化处理工具启动")

    # 如果没有输入文件，自动生成演示数据
    input_path = Path(INPUT_DIR)
    if not input_path.exists() or not any(input_path.iterdir()):
        create_demo_data()

    processor = ExcelProcessor()

    # 批量处理
    df = processor.process_all()

    if df.empty:
        return

    # 生成报告
    report = processor.generate_report(df)
    logger.info(f"\n{'='*50}")
    logger.info("📊 汇总报告")
    logger.info(f"   总记录: {report.get('总记录数', 0)}")
    logger.info(f"   文件数: {report.get('文件数', 0)}")
    for col, stats in report.get("数值列统计", {}).items():
        logger.info(f"   [{col}] 总和={stats['总和']}, 均值={stats['平均值']}")
    logger.info(f"{'='*50}")

    # 导出
    processor.save_results(df)
    logger.info("🎉 处理完成！请查看 output/ 文件夹")


if __name__ == "__main__":
    main()
