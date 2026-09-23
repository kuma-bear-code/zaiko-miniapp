# Daily LINE digest

## Data flow

1. Apps Script refreshes `商品分析` from the current inventory and consumption source at 5:00-6:00 Japan time.
2. A ChatGPT scheduled run at 6:00 checks yesterday's rows in `ConsumptionLog`, current stock in `Inventory`, and GAS-generated purchase quantities in `商品分析`.
3. It writes one Japanese digest row per date to `LINE配信` with `対象日`, `配信文`, `配信状態`, `配信日時`, and `エラー` columns. New rows start as `未送信`.
4. The Apps Script trigger sends the previous day's `未送信` row to the existing `GROUP_ID` at 7:00-8:00 Japan time, then records `配信済み` and the delivery timestamp.

The scheduled summary must include every consumption row, including same-day rows with the same item and quantity. Those entries represent real use and count toward forecasts. When yesterday has no rows, the summary should say so and use the latest recorded date as context. It must copy purchase quantities from `商品分析` rather than calculate them independently.

The ChatGPT job should upsert by `対象日`, preserve rows already marked `配信済み`, and write `未送信` only for a new, undelivered summary. It should not edit `Inventory`, `ConsumptionLog`, forecast formulas, or any existing analysis tabs.

## Apps Script setup

After deploying `Code.gs`, run `setupDailyLineDigestTrigger()` once in Apps Script and grant the requested spreadsheet and external request permissions. It creates daily triggers for `refreshDailyLineDigestSource` at 5:00 and `sendPendingLineDigest` at 7:00 in `Asia/Tokyo`; repeated setup calls do not create duplicate triggers.

LINE delivery failures are recorded as `配信失敗` with the API error in `エラー`. Review the message and error before changing the row back to `未送信` for a manual retry. A row marked `配信中` is intentionally skipped to avoid an automatic duplicate if an execution ended while LINE's response was uncertain.

## Scheduled ChatGPT prompt

Run daily at 6:00 Japan time. Read the Google Sheet `在庫管理bot` (ID `1eOg7cd1V6BlvsbiJgeoHXAty1TwLOwRXUEpvyhIBtII`). For yesterday in `Asia/Tokyo`, summarize all `ConsumptionLog` entries without deduplicating them, then summarize purchase candidates from `商品分析` using its existing status and `購入推奨数`. Write a concise, natural Japanese LINE message to one row of `LINE配信`, using yesterday's date as `対象日` in `yyyy-MM-dd` format and `未送信` as `配信状態`. Include `配信文` as the complete message. If the row for that date is already `配信済み`, leave it unchanged. If there were no consumption entries, say that clearly and report the latest date with recorded use. Never edit the source inventory or consumption data, and never recalculate purchase quantities.
