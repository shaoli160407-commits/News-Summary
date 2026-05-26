# DeepSeek API Notes

Use official DeepSeek docs as the source of truth before changing API code.

Current assumptions checked on 2026-05-26:

- OpenAI-compatible base URL: `https://api.deepseek.com`
- Anthropic-compatible base URL: `https://api.deepseek.com/anthropic`
- Recommended models include `deepseek-v4-pro` and `deepseek-v4-flash`
- The Anthropic-compatible API supports `tools`, `tool_choice`, `server_tool_use`, and `web_search_tool_result`
- DeepSeek's Claude Code integration documentation says Web Search can be invoked through the API when the model determines search is needed

For raw script calls, use:

- URL: `https://api.deepseek.com/anthropic/v1/messages`
- Header: `x-api-key: <DEEPSEEK_API_KEY>`
- Header: `anthropic-version: 2023-06-01`
- JSON body with `model`, `max_tokens`, `system`, `messages`, and web-search tool configuration.

If the web-search tool shape changes, update only `call_deepseek_research()` in `scripts/build_trade_daily.py`.
