# pull-asset-prices

A script for Google Sheets that automates pulling asset prices for Stocks, Crypto, and Metals.

## 📄 Disclaimer & Usage Policy

This tool/script is intended for **personal use only** and is provided **as-is**, without warranty of any kind. By using this tool, you acknowledge and agree to the following:

---

### 🔍 Data Providers

This tool retrieves publicly available data from the following sources:

#### 🟦 Yahoo Finance
- **Endpoint**: `https://query1.finance.yahoo.com`
- Yahoo's finance data is accessed via internal endpoints not officially documented or supported. Use of Yahoo's data is subject to their [Terms of Service](https://legal.yahoo.com/us/en/yahoo/terms/otos/index.html), which **prohibit automated access** and commercial reuse without permission.
- This tool is **not affiliated with, endorsed by, or officially supported by Yahoo**.

#### 🟧 KuCoin
- **Endpoint**: `https://api.kucoin.com/api/v1/market/allTickers`
- This endpoint is publicly accessible and does **not require authentication**. However, usage is still governed by KuCoin’s [Terms of Use](https://www.kucoin.com/legal/terms-of-use). This script respects KuCoin’s rate limits and is intended for **non-commercial, personal portfolio tracking** only.
- This tool is **not affiliated with or endorsed by KuCoin**.

#### 🟨 Gold-API
- **Website**: [https://gold-api.com](https://gold-api.com)
- Gold-API provides access to gold and precious metals prices. This tool may optionally use the Gold-API free or paid tiers depending on user configuration. Use of Gold-API is governed by their [Terms of Use](https://gold-api.com/terms).
- Gold-API access in this tool requires a **user-supplied API key** and is subject to their usage limits and licensing terms.
- This tool is **not affiliated with or endorsed by Gold-API**.

---

### 📌 Usage Restrictions

- **Do not use this tool for commercial purposes** unless you have the appropriate licensing or permission from the data providers listed above.
- **Do not modify this tool to scrape, redistribute, or expose raw API data** in public-facing apps, dashboards, or resellable products.
- **Respect API rate limits and fair use** policies of each provider to avoid throttling or access blocks.

---

### 💬 Transparency

This tool was created for educational and personal portfolio tracking purposes.  
If you are the owner of one of the services above and have concerns about this tool, please contact the maintainer to discuss proper usage or removal.

---

### ☕ Support

This tool is free to use. If you find it helpful and want to support its development, you can [donate here](#your-ko-fi-or-donation-link).
