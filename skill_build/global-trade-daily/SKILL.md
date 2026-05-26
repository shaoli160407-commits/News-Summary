---
name: global-trade-daily
description: Generate a global foreign trade news daily report with DeepSeek API research and summarization. Use when the user asks for 外贸资讯日报, DeepSeek-powered global trade daily/news brief, import-export market updates, tariff/trade sanctions summaries, supply chain trade reports, or to create and email a DOCX report about recent global trade news.
---

# Global Trade Daily

Use DeepSeek API to research and summarize recent global trade news, then create a Chinese-first DOCX report and optionally email it to the user's own Gmail.

## Configuration

Use a workspace `.env` file. Never print secret values.

Required for DeepSeek research:

```env
DEEPSEEK_API_KEY=...
```

Required only when sending email by SMTP:

```env
GMAIL_ADDRESS=...
GMAIL_APP_PASSWORD=...
```

DeepSeek notes:

- Use the official DeepSeek Anthropic-compatible API for web-search research: `https://api.deepseek.com/anthropic`.
- Use model `deepseek-v4-pro` by default.
- The script uses the Anthropic Messages shape with the server web-search tool. If DeepSeek changes that compatibility surface, check `references/deepseek_api.md` and update the script.

## Workflow

1. **Check runtime prerequisites**
   - Use the active Python environment.
   - Ensure `python-docx` and `python-dotenv` are installed.
   - Do not require Codex web browsing for the report itself; the script should call DeepSeek API for research and summarization.

2. **Run DeepSeek research**
   - Run `scripts/build_trade_daily.py --deepseek-research --output-dir <output_dir>`.
   - Use a 5-day news window unless the user specifies another range.
   - Search intent must cover:
     - `全球外贸 最新动态`, `进出口政策`, `关税 贸易`, `供应链 外贸`, `汇率 出口`
     - `global trade news`, `import export policy`, `tariffs trade war`, `supply chain disruption`, `trade sanctions`
   - Require at least 10 source-backed items with title, source, date, URL, summary, and trade impact.
   - Prioritize tariff policy, sanctions, import/export data, exchange rates, commodities, logistics disruptions, and trade agreements.

3. **Generate report content**
   - DeepSeek must output JSON matching `references/payload_schema.md`.
   - Executive summary: 200 Chinese characters or fewer.
   - Categories: `政策法规`, `市场行情`, `供应链与物流`, `汇率与金融`.
   - Include 3-5 items per category when available.
   - Each item must include title, date, 3-5 sentence core content, and potential foreign trade impact.
   - Market trend conclusion: 100 Chinese characters or fewer.

4. **Build DOCX**
   - The script converts DeepSeek JSON into `output/外贸资讯日报_YYYY-MM-DD.docx`.
   - Source URLs must be retained in the source list and per-item source note.

5. **Send email**
   - Add `--send-email` when the user wants SMTP delivery.
   - Recipient equals sender.
   - Email subject defaults to `【外贸日报】YYYY年MM月DD日 全球外贸资讯总结`.
   - On sandbox network failure during DeepSeek or SMTP calls, rerun with escalation.

6. **Verify**
   - Confirm the DOCX exists and can be opened with `python-docx`.
   - If the documents skill and LibreOffice/soffice are available, render to PNG and inspect pages before final delivery.
   - If rendering is unavailable because LibreOffice/soffice is missing, report that visual QA was skipped and include structural verification.

## Script Usage

DeepSeek researches, summarizes, builds DOCX, and sends email:

```powershell
python scripts/build_trade_daily.py --deepseek-research --output-dir .\output --send-email
```

DeepSeek researches and builds DOCX without email:

```powershell
python scripts/build_trade_daily.py --deepseek-research --output-dir .\output
```

Build DOCX from an already prepared JSON payload:

```powershell
python scripts/build_trade_daily.py .\trade_daily_payload.json --output-dir .\output
```

The script prints Chinese progress messages and raises clear errors for missing payload fields, missing credentials, DeepSeek API failures, DOCX generation failures, or SMTP failures.
