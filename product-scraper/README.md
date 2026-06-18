# 🕷️ Product Scraper — 通用电商爬虫

> Python 爬虫演示项目 | 适用于 Fiverr 展示

## 功能亮点

- ✅ **多格式导出**：同时输出 CSV / JSON / Excel
- ✅ **健壮错误处理**：自动重试（指数退避）、跳过损坏数据
- ✅ **礼貌爬取**：自定义请求间隔，避免被封
- ✅ **分页自动处理**：自动遍历所有页面，也可限制页数
- ✅ **数据统计**：爬取完成后自动输出价格汇总、评分分布
- ✅ **类型安全**：使用 dataclass + 类型注解

## 技术栈

| 技术 | 用途 |
|------|------|
| `requests` | HTTP 请求 |
| `BeautifulSoup4` | HTML 解析 |
| `pandas` | 数据分析 + Excel 导出 |
| `lxml` | 高性能 HTML 解析器 |

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行（自动爬取 + 导出 3 种格式）
python scraper.py
```

## 输出示例

```
🚀 产品爬虫启动
2024-01-15 20:30:01 [INFO] Fetching: https://books.toscrape.com/...
2024-01-15 20:30:02 [INFO] ✓ 第 1 页完成，共爬取 20 个产品
...
==================================================
📊 爬取摘要
   总产品数: 1000
   总价值:   £35142.15
   均价:     £35.14
   最高价:   £57.25
   评分分布:
   Three    320
   Four     246
   One      226
   Two      124
   Five      84
==================================================
✓ 已保存 CSV: products.csv (1000 条)
✓ 已保存 JSON: products.json (1000 条)
✓ 已保存 Excel: products.xlsx (1000 条)
🎉 完成！
```

## 改造指南（适配不同网站）

在 `_parse_product()` 方法中修改 CSS 选择器即可：

```python
# 改之前（Books to Scrape）
title = article.find("h3").find("a")["title"]
price = article.find("p", class_="price_color").text

# 改之后（你的目标网站）
title = article.find("h2", class_="product-name").text
price = article.find("span", class_="sale-price").text
```

## 目录结构

```
product-scraper/
├── scraper.py          # 主程序
├── requirements.txt    # 依赖
└── README.md          # 本文档
```

---

👨‍💻 **Fiverr 接单时**：将此项目作为案例发给客户，证明你能专业地爬取和清洗数据。
