from CitationUtility import CitationUtility
dir_base = r"U:\MDD-CTRL\T0040--Aptasensors-Book-Chapter"

doc_file = f"{dir_base}/T0040--Chapter-8--MicroFluidics.docx"
dir_ref = f"{dir_base}/references"
_params = {
  "base_path": dir_base,
  "dir_references": dir_ref
}

_cu = CitationUtility(**_params)
_cu.format_references(doc_file)
# _cu.ref_mapping
