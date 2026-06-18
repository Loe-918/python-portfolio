"""
Product Scraper — 通用电商产品信息爬虫
===========================================
演示用途：爬取 Books to Scrape（练习用电商网站）
实际接单时可以快速改造成爬取真实目标网站。

作者：Liam (Fiverr Portfolio)
"""

import csv
import json
import logging
import time
from dataclasses import dataclass, asdict
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============== 配置 ==============
BASE_URL = "https://books.toscrape.com/catalogue/"
START_URL = f"{BASE_URL}page-1.html"
REQUEST_DELAY = 1.0  # 请求间隔（秒），礼貌爬取
OUTPUT_CSV = "products.csv"
OUTPUT_JSON = "products.json"
OUTPUT_EXCEL = "products.xlsx"
MAX_PAGES: Optional[int] = None  # None = 所有页面；设为数字限制页数

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class Product:
    """产品数据结构"""
    title: str
    price: str          # 保留原始格式：£51.77
    price_value: float  # 纯数字：51.77
    availability: str
    rating: str         # One / Two / Three / Four / Five
    category: str
    product_url: str
    image_url: str


class BooksScraper:
    """优雅的爬虫类 — 专业的错误处理、日志、容错机制"""

    def __init__(self, base_url: str = START_URL, delay: float = REQUEST_DELAY):
        self.base_url = base_url
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        })
        self.products: list[Product] = []

    # ---------- 网络层 ----------
    def _get(self, url: str, retries: int = 3) -> Optional[BeautifulSoup]:
        """带重试的 GET 请求"""
        for attempt in range(1, retries + 1):
            try:
                logger.info(f"Fetching: {url}")
                resp = self.session.get(url, timeout=30)
                resp.raise_for_status()
                return BeautifulSoup(resp.text, "lxml")
            except requests.RequestException as e:
                logger.warning(f"请求失败 (尝试 {attempt}/{retries}): {e}")
                if attempt < retries:
                    time.sleep(2 ** attempt)  # 指数退避
        logger.error(f"放弃请求: {url}")
        return None

    # ---------- 解析层 ----------
    def _parse_product(self, article) -> Optional[Product]:
        """解析单个产品卡片"""
        try:
            # 图片 URL
            img = article.find("img")
            image_url = img["src"].replace("../..", "https://books.toscrape.com") if img else ""

            # 标题 + 详情链接
            title_tag = article.find("h3").find("a") if article.find("h3") else None
            title = title_tag.get("title", "No Title") if title_tag else "No Title"
            product_url = (
                BASE_URL + title_tag["href"] if title_tag and title_tag.get("href") else ""
            )

            # 价格
            price_tag = article.find("p", class_="price_color")
            price = price_tag.text.strip() if price_tag else "£0.00"
            try:
                price_value = float(price.replace("£", "").replace("Â", "").strip())
            except ValueError:
                price_value = 0.0

            # 库存
            avail_tag = article.find("p", class_="instock availability")
            availability = avail_tag.text.strip() if avail_tag else "Unknown"

            # 评分（class 是 "star-rating One" 这种格式）
            rating_tag = article.find("p", class_="star-rating")
            rating = "No Rating"
            if rating_tag:
                classes = rating_tag.get("class", [])
                for cls in classes:
                    if cls in ("One", "Two", "Three", "Four", "Five"):
                        rating = cls
                        break

            return Product(
                title=title,
                price=price,
                price_value=price_value,
                availability=availability,
                rating=rating,
                category="Books",
                product_url=product_url,
                image_url=image_url,
            )
        except Exception as e:
            logger.warning(f"解析产品卡片失败: {e}")
            return None

    def _get_category_from_page(self, soup: BeautifulSoup) -> str:
        """从页面提取分类名"""
        breadcrumb = soup.find("ul", class_="breadcrumb")
        if breadcrumb:
            items = breadcrumb.find_all("li")
            return items[-1].text.strip() if items else "Unknown"
        return "Unknown"

    # ---------- 主逻辑 ----------
    def scrape(self, max_pages: Optional[int] = None) -> list[Product]:
        """执行爬取：遍历分页，解析每页产品"""
        page = 1
        while True:
            if max_pages and page > max_pages:
                break

            url = f"{BASE_URL}page-{page}.html"
            soup = self._get(url)

            if not soup:
                logger.info(f"页面 {page} 无响应，停止翻页。")
                break

            # 检查是否有内容
            articles = soup.find_all("article", class_="product_pod")
            if not articles:
                logger.info(f"页面 {page} 无产品，翻页结束。")
                break

            for article in articles:
                product = self._parse_product(article)
                if product:
                    self.products.append(product)

            logger.info(f"✓ 第 {page} 页完成，共爬取 {len(self.products)} 个产品")
            page += 1
            time.sleep(self.delay)

        return self.products

    # ---------- 导出层 ----------
    def to_dataframe(self) -> pd.DataFrame:
        """转为 pandas DataFrame，方便数据分析和 Excel 导出"""
        return pd.DataFrame([asdict(p) for p in self.products])

    def save_all(self, csv_path: str = OUTPUT_CSV, json_path: str = OUTPUT_JSON,
                 excel_path: str = OUTPUT_EXCEL):
        """一键保存为 CSV + JSON + Excel"""
        df = self.to_dataframe()

        # CSV
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        logger.info(f"✓ 已保存 CSV: {csv_path} ({len(df)} 条)")

        # JSON
        df.to_json(json_path, orient="records", force_ascii=False, indent=2)
        logger.info(f"✓ 已保存 JSON: {json_path} ({len(df)} 条)")

        # Excel
        df.to_excel(excel_path, index=False, engine="openpyxl")
        logger.info(f"✓ 已保存 Excel: {excel_path} ({len(df)} 条)")


# ============== 入口 ==============
def main():
    logger.info("🚀 产品爬虫启动")
    scraper = BooksScraper()

    # 爬取（不设 max_pages 则爬全部）
    products = scraper.scrape(max_pages=MAX_PAGES)

    if not products:
        logger.warning("⚠️ 未爬取到任何产品")
        return

    # 统计摘要
    df = scraper.to_dataframe()
    logger.info(f"\n{'='*50}")
    logger.info(f"📊 爬取摘要")
    logger.info(f"   总产品数: {len(products)}")
    logger.info(f"   总价值:   £{df['price_value'].sum():.2f}")
    logger.info(f"   均价:     £{df['price_value'].mean():.2f}")
    logger.info(f"   最高价:   £{df['price_value'].max():.2f}")
    logger.info(f"   评分分布:\n{df['rating'].value_counts().to_string()}")
    logger.info(f"{'='*50}")

    # 导出
    scraper.save_all()
    logger.info("🎉 完成！")


if __name__ == "__main__":
    main()
