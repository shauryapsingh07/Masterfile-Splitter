# Masterfile Splitter

Splits a marketplace masterfile (.xlsx) into **Complete** and **Incomplete** SKUs. It's been built for the Mirakl-style Tesco UK and Macy's templates. It runs as a Streamlit web app (`app.py`) or from the command line (`split_masterfile.py`). The splitting engine uses only the Python standard library.

## How it decides

- **Mandatory means red.** The templates paint a cell red when that attribute is mandatory for the SKU's category. They do this with a conditional-formatting rule that looks up the category in the `Columns` sheet. The tool reads that rule from the workbook, so it marks exactly the cells Excel shows in red, row by row.
- **Complete** means every red cell in the row has something in it. **Incomplete** means at least one red cell is blank.
- **Cell values are never judged or changed.** Any non-blank entry (including "NA") counts as filled. Yellow (recommended) and gray (NA) cells are ignored.
- **Category problems.** If a row's category is blank, or isn't listed in the `Columns` sheet, the row goes to Incomplete with Category Code listed as missing.

## Output (per masterfile)

| File | Contents |
|---|---|
| `<name> - Complete.xlsx` | Complete SKUs in the original template, ready to upload |
| `<name> - Incomplete.xlsx` | Incomplete SKUs in the original template, with a **Missing Attributes** column added after the last column |
| `<name> - Summary.csv` | Total, Complete and Incomplete counts, plus every missing attribute ranked by how many SKUs miss it |

**What stays unchanged:** every sheet, style, dropdown, colour rule, named range and macro is copied as-is. Only the SKU rows on the `Data` sheet are selected and renumbered from row 3. The `Error Details` sheet is copied unchanged, so its line numbers still refer to the original file. Completely blank rows are skipped and counted in the summary.

## Deploy on Streamlit Community Cloud (via GitHub)

1. **Create the repo.** On GitHub, create a new repository, for example `masterfile-splitter`. Make it private if your plan allows.
2. **Upload the files.** Click **Add file → Upload files** and drag in the contents of this folder: `app.py`, `split_masterfile.py`, `requirements.txt`, `README.md` and `.gitignore`. Commit to `main`. Don't upload any catalog `.xlsx` files.
3. **Create the app.** Go to <https://share.streamlit.io>, sign in with GitHub, and click **Create app → Deploy a public app from GitHub**.
4. **Point it at the repo.** Choose your repository, branch `main`, and main file path `app.py`. Under **Advanced settings**, pick Python 3.11 or 3.12.
5. **Deploy.** Click **Deploy**. The first build takes a minute or two. You'll get a URL like `https://<your-app>.streamlit.app`.
6. **Restrict who can open it.** In the app's **Settings → Sharing**, limit viewers to specific emails, because people will upload catalog data.

Any later commit to `main` redeploys the app automatically.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Command line

```bash
python split_masterfile.py "Tesco UK All Catalog.xlsx"
python split_masterfile.py "Tesco UK All Catalog.xlsx" "Macy's All Catalog.xlsx" -o split
python split_masterfile.py -l files.txt
```

| Option | Description |
|---|---|
| `files` | One or more masterfile paths |
| `-l`, `--list-file` | Text file with one masterfile path per line (`#` lines ignored) |
| `-o`, `--out-dir` | Output folder (default `split_output`) |
| `--top` | How many top missing attributes to print (default 10) |
