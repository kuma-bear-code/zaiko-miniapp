# GAS / LIFF release checklist

The Git branch and Draft PR do not deploy `Code.gs` or GitHub Pages. Keep `main`
and the current GAS deployment active while reviewing this branch.

## Before deployment

1. Make a copy of the spreadsheet and record the current GAS deployment version.
2. Confirm Script Properties include `SPREADSHEET_ID`, `CHANNEL_ACCESS_TOKEN`,
   `LIFF_CHANNEL_ID`, and `ALLOWED_LINE_USER_IDS`. Do not commit their values.
3. Deploy the branch's `Code.gs` to a test GAS project pointing at the copied
   spreadsheet. Test the branch's `index.html` and `shopping.html` against it.
4. Run `node forecast.test.js`. Run `refreshForecastAnalysis()` once in the test
   GAS project. Check `Settings`, `商品設定`, `商品分析`, `異常候補`, and `Dashboard`.

## Test with the copied spreadsheet

- Confirm Inventory and ConsumptionLog row counts and a sample of original rows
  match the copy before making test changes.
- Check stock 0, below minimum, exactly minimum, and above minimum.
- Check 0, 1, and 2+ consumption records, plus 180-day, 365-day, and all-time
  forecast bases. Check a long-unused item, a duplicate record, and an outlier.
- Check pack sizes 1 and greater than 1, including a product rename.
- From Web, add, edit location, adjust up/down, delete, and undo a product.
  Confirm the corresponding Sheet values after reloading Web. Confirm deleting
  the product leaves its ConsumptionLog rows intact.
- From shopping.html, restock multiple checked items. Verify the exact Sheet
  stock values and the reloaded Web values before accepting success.
- From LINE, buy and consume a test product, then reload Web and the Sheet.
- Change a Sheet stock cell directly, then confirm Web reload picks it up and
  rejects an edit based on the stale stock value.
- Check that a request without a LINE ID token cannot read or mutate Web data.

## Before production switch

Review the test results, API compatibility, and the LINE webhook authentication
boundary. Deploy GAS and the Web/LIFF pages together, then repeat the readback
checks against a dedicated test product in the production spreadsheet. Keep the
previous GAS deployment version available for rollback.
