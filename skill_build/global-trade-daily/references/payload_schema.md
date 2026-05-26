# Payload Schema

Use UTF-8 JSON. Required top-level fields:

```json
{
  "report_date": "YYYY-MM-DD",
  "title": "全球外贸资讯日报",
  "executive_summary": "200字以内中文摘要",
  "trend": "100字以内中文趋势研判",
  "categories": [
    {
      "name": "政策法规",
      "items": [
        {
          "title": "新闻标题",
          "source": "来源名称",
          "date": "YYYY-MM-DD",
          "url": "https://source.example/article",
          "summary": "3-5句话核心内容",
          "impact": "对外贸的潜在影响"
        }
      ]
    }
  ]
}
```

Recommended category names:

- `政策法规`
- `市场行情`
- `供应链与物流`
- `汇率与金融`

Optional top-level field:

```json
{
  "email_subject": "【外贸日报】YYYY年MM月DD日 全球外贸资讯总结"
}
```

If `email_subject` is omitted, the script generates it from `report_date`.
