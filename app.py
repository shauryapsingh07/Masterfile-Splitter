"""Streamlit front end for the Masterfile Splitter.

Run locally:   streamlit run app.py
Deploy:        push this folder to GitHub, then create an app on Streamlit Community Cloud
               with app.py as the main file.
"""
import io
import zipfile

import streamlit as st

from split_masterfile import output_names, split_workbook

st.set_page_config(page_title="Masterfile Splitter", page_icon="📑", layout="wide")


@st.cache_data(show_spinner=False, max_entries=20)
def run_split(data: bytes, name: str):
    return split_workbook(io.BytesIO(data), name)


st.title("Masterfile Splitter")
st.write(
    "Upload one or more masterfiles. Each SKU is checked against the cells the template "
    "highlights in **red** (mandatory for that SKU's category). A SKU is **Complete** when "
    "every red cell has a value. Values are never judged or changed, and any non-blank entry counts as filled."
)

files = st.file_uploader("Masterfiles (.xlsx)", type=["xlsx", "xlsm"], accept_multiple_files=True)

if not files:
    st.info("Waiting for files. Your template structure, formatting, dropdowns and macros are kept unchanged.")
    st.stop()

results = []
for f in files:
    with st.spinner(f"Splitting {f.name} ..."):
        try:
            results.append(run_split(f.getvalue(), f.name))
        except Exception as e:  # show the problem, keep going with the other files
            st.error(f"**{f.name}** could not be split: {e}")

if not results:
    st.stop()

# One zip with every output, for convenience.
bundle = io.BytesIO()
with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
    for res in results:
        for key, filename in output_names(res).items():
            z.writestr(filename, res[key])

total = sum(r["total"] for r in results)
st.download_button(
    f"Download all outputs ({len(results)} file{'s' if len(results) > 1 else ''}, {total} SKUs) as .zip",
    bundle.getvalue(),
    file_name="masterfile_split_outputs.zip",
    mime="application/zip",
    type="primary",
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

for res in results:
    st.divider()
    st.subheader(res["name"])
    names = output_names(res)
    pct = (lambda n: f"{100.0 * n / res['total']:.1f}%") if res["total"] else (lambda n: "0%")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total SKUs", f"{res['total']:,}")
    c2.metric("Complete", f"{res['complete']:,}", pct(res["complete"]), delta_color="off")
    c3.metric("Incomplete", f"{res['incomplete']:,}", pct(res["incomplete"]), delta_color="off")
    if res["blank_rows"]:
        st.caption(f"{res['blank_rows']} completely blank row(s) were skipped.")

    d1, d2, d3 = st.columns(3)
    d1.download_button("Complete SKUs (.xlsx)", res["complete_xlsx"], file_name=names["complete_xlsx"],
                       mime=XLSX_MIME, key=f"c-{res['name']}", use_container_width=True)
    d2.download_button("Incomplete SKUs (.xlsx)", res["incomplete_xlsx"], file_name=names["incomplete_xlsx"],
                       mime=XLSX_MIME, key=f"i-{res['name']}", use_container_width=True)
    d3.download_button("Summary report (.csv)", res["summary_csv"], file_name=names["summary_csv"],
                       mime="text/csv", key=f"s-{res['name']}", use_container_width=True)

    if res["missing_counts"]:
        st.markdown("**Most commonly missing attributes**")
        rows = [{"Attribute": m["attribute"], "Column": m["column"], "Code": m["code"],
                 "SKUs missing": m["skus"],
                 "% of SKUs": round(100.0 * m["skus"] / res["total"], 1) if res["total"] else 0.0}
                for m in res["missing_counts"]]
        st.dataframe(
            rows,
            hide_index=True,
            use_container_width=True,
            column_config={"% of SKUs": st.column_config.ProgressColumn(
                "% of SKUs", min_value=0, max_value=100, format="%.1f%%")},
        )
    else:
        st.success("Every SKU has all of its red (mandatory) cells filled.")
