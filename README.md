# pull-asset-prices

A script for Google Sheets that automates pulling asset prices for Stocks, Crypto, and Metals.

## ✅ What This Script Does (Summary)

This Google Apps Script powers a **Google Sheets tool called "Pull Asset Prices"**. Here's what it does:

### 🌐 Fetches asset prices from:
- **Yahoo Finance** → Stock/ETF prices via an *unofficial* endpoint  
- **KuCoin** → Crypto prices (BTC, ETH, etc.)  
- **Gold-API** → Precious metal prices (Gold & Silver)  

### 📋 Core Features:
- Builds a custom "Asset Prices" tab in your sheet  
- Lets you enter tickers for stocks and crypto  
- Automatically fetches and updates their live prices  
- Creates named ranges (like `=BTC_Price`) so you can reference prices in formulas  
- Includes a custom menu in the spreadsheet UI:
  - “Pull Prices Now”
  - “Price Variable - Usage Example”
  - “Check for Updates”
- Automatically updates prices **hourly** via a time-based trigger  
- Logs any update errors or warnings to a visible section in the sheet  

---

## 🔐 Is It Safe to Run?

### ✅ Yes — with a few considerations:

#### The Good:
- No destructive actions (doesn't delete your data, doesn't send data elsewhere)  
- Uses only public or semi-public APIs (no authentication or private info involved)  
- Clearly documented, with thoughtful error handling and logging  
- Source is visible and editable by you  

#### Minor Considerations:
- **Uses unofficial Yahoo Finance endpoint**  
  - Could break if Yahoo changes it  
  - Might technically violate TOS if abused  
- **Fetches data from external URLs**  
  - All are reputable (Yahoo, KuCoin, Gold-API)  
- **Automatically installs a trigger**  
  - Runs hourly; you can remove this if needed  

---

## 💬 TL;DR

**Yes, it’s safe to run.** It’s well-written, nicely documented, and behaves like a solid personal-use Sheets tool.  
Just keep in mind:
- Unofficial Yahoo API
- External URL fetches
- Hourly update trigger



# 📄 Disclaimer & Usage Policy

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
