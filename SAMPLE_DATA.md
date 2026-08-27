# Where the data lives

**Not in this repo, by design.**

The app takes everything through the browser:

| | How it arrives |
|---|---|
| Master mapping table | uploaded once per session |
| Uniware reports | uploaded (or fetched, once that is enabled) |
| Amazon Vendor Central | uploaded each cycle |
| Retail store files | uploaded each cycle |

Fastest way to load a whole cycle: zip the input folder so it looks like this,
then use **Upload everything at once (.zip)** in the sidebar.

```
domin8_inputs.zip
  Marketplace product id Master.xlsx
  uniware/         Tally GST, Tally Return GST, Inventory Snapshot,
                   Item Master, Purchase Orders
  amazon vc/       Sales_ASIN_*.csv, Inventory_ASIN_*.csv
  retail stores/   each store's sale + stock-on-hand files
```

Folder names containing `uniware`, `amazon` or `retail`/`store` are routed
automatically. Anything with a `Sku Code` column is treated as the master table.

## Running locally against a folder instead

If `reports/input/` exists beside the code, the app seeds from it on startup so
you can work without uploading. That folder is gitignored. `run_pipeline.py`
still works exactly as before for CLI runs.
