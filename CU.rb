import sys as SYSTEM
SYSTEM.path.append(r"D:\Documents\R\RProjects\DevWork\Apps\UtilityLib2")

import re as REGEX
import pandas as PD
import json as JSON
from zipfile import ZipFile
from random import randint as RandInt

from lxml import etree as ET, objectify as Objectify
from pylatexenc.latex2text import LatexNodes2Text
from docx import Document

from UtilityManager import UtilityManager

"""

* [Render Latex text with python](https://stackoverflow.com/a/38168602/6213452)
* Fixed DOI missing issue

"""


class CitationUtility:
  def __init__(self, *args, **kwargs):
    self.__update_attr(**kwargs)
    self.sources_add_root()

  def __update_attr(self, *args, **kwargs):
    if not hasattr(self, "__defaults"): self.__defaults =  {
        "debug": False,
        "sources_xml": None,
        "Document": Document,
        "XMLBuilder": ET,
        "ZipFile": ZipFile,
        "base_path": None,
        "sources_filename": "UtilitySources.xml",
        "sources_filepath": None,
        "sources_namespaces": {"b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"},
        "references": [],
        "citation_xml_string": """<w:sdt><w:sdtPr><w:rPr><w:color w:val="000000"/><w:szCs w:val="24"/></w:rPr><w:id w:val="{citation_id}"/><w:citation/></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:r w:rsidR="{citation_hex}"><w:rPr><w:color w:val="000000"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r w:rsidR="{citation_hex}"><w:rPr><w:color w:val="000000"/><w:szCs w:val="24"/><w:lang w:val="en-IN"/></w:rPr><w:instrText xml:space="preserve"> CITATION {citation_placeholder} \l 16393 </w:instrText></w:r><w:r w:rsidR="{citation_hex}"><w:rPr><w:color w:val="000000"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="0071451E" w:rsidRPr="0071451E"><w:rPr><w:noProof/><w:color w:val="000000"/><w:szCs w:val="24"/><w:lang w:val="en-IN"/></w:rPr><w:t>({citation_placeholder})</w:t></w:r><w:r w:rsidR="{citation_hex}"><w:rPr><w:color w:val="000000"/><w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:sdtContent></w:sdt>""",
        "utility": UtilityManager(db_path = "test-references.db"),
        "endpoint_doi_conversion": 'http://dx.doi.org',
        "endpoint_pmid_conversion": "https://api.paperpile.com/api/public/convert",
        "regex_parse_bibtext": REGEX.compile(r'(?P<key>.*)\s{0,}=\s{0,}(?P<value>.*)'),
        "regex_remove_tags": REGEX.compile('<.*?>'),
        "regex_find_ref": r'\[(.*?)\]',
        "regex_validate_ref_text": r'(doi|pmid|pmc)',
        "ref_discovered": [],
        "ref_mapping": [],
        "ref_details": None,
        "dir_references": None,
      }

    # Set all defaults
    [setattr(self, _k, self.__defaults[_k]) for _k in self.__defaults.keys() if not hasattr(self, _k)]
    self.__defaults = dict() # Unset defaults to prevent running for second time
    [setattr(self, _k, kwargs[_k]) for _k in kwargs.keys()]

    if self.sources_filepath is None and self.base_path is not None:
      self.sources_filepath = self.base_path

  def sources_add_root(self, *args, **kwargs):
    self.sources_xml = self.XMLBuilder.Element(
        self.XMLBuilder.QName(self.sources_namespaces["b"], 'Sources'),
                        nsmap = self.sources_namespaces)
    self.sources_xml.attrib['xmlns'] = self.sources_namespaces['b']

  def __sources_add_authors(self, *args, **kwargs):
    _element = args[0] if len(args) > 0 else kwargs.get("element")
    _authors = args[1] if len(args) > 1 else kwargs.get("authors", [])

    if isinstance(_authors, str):
      _authors = [_authors]

    if len(_authors) > 0 and _element is not None:
      _author = self.XMLBuilder.SubElement(_element, self.XMLBuilder.QName(self.sources_namespaces['b'], "Author"))
      _author = self.XMLBuilder.SubElement(_author, self.XMLBuilder.QName(self.sources_namespaces['b'], "Author"))
      _name_list = self.XMLBuilder.SubElement(_author, self.XMLBuilder.QName(self.sources_namespaces['b'], "NameList"))

      for _author in _authors:
        # Loop for person
        _person = self.XMLBuilder.SubElement(_name_list, self.XMLBuilder.QName(self.sources_namespaces['b'], "Person"))
        _last = self.XMLBuilder.SubElement(_person, self.XMLBuilder.QName(self.sources_namespaces['b'], "Last"))
        _middle = self.XMLBuilder.SubElement(_person, self.XMLBuilder.QName(self.sources_namespaces['b'], "Middle"))
        _first = self.XMLBuilder.SubElement(_person, self.XMLBuilder.QName(self.sources_namespaces['b'], "First"))

        if isinstance(_author, str):
          _author = [_author]

        _last.text = _author[0] if len(_author) > 0 else ""
        _first.text = _author[1] if len(_author) > 1 else ""
        _middle.text = _author[2] if len(_author) > 2 else ""

    return _element

  def sources_add_article(self, *args, **kwargs):
    """
    Create a XML element
    """
    _data = args[0] if len(args) > 0 else kwargs.get("data", {})

    _article = self.XMLBuilder.SubElement(self.sources_xml, self.XMLBuilder.QName(self.sources_namespaces['b'], 'Source'))

    Tag = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Tag"))
    Issue = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Issue"))
    Year = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Year"))
    Volume = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Volume"))
    SourceType = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "SourceType"))
    Title = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Title"))
    DOI = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "DOI"))

    self.__sources_add_authors(_article, _data.get("Author", []))

    Pages = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Pages"))
    Month = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "Month"))
    JournalName = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "JournalName"))
    City = self.XMLBuilder.SubElement(_article, self.XMLBuilder.QName(self.sources_namespaces['b'], "City"))

    Tag.text = _data.get("Tag")
    Issue.text = _data.get("Issue")
    Year.text = _data.get("Year")
    Volume.text = _data.get("Volume")
    SourceType.text = _data.get("SourceType")
    Title.text = _data.get("Title")
    DOI.text = _data.get("DOI")
    Pages.text = _data.get("Pages")
    Month.text = _data.get("Month")
    JournalName.text = _data.get("JournalName")
    City.text = _data.get("City")

  def sources_write_xml(self, *args, **kwargs):
    _sources_filename = args[0] if len(args) > 0 else kwargs.get("sources_filename", getattr(self, "sources_filename"))
    _sources_xml = args[1] if len(args) > 1 else kwargs.get("sources_xml", getattr(self, "sources_xml"))
    _sources_filepath = args[2] if len(args) > 2 else kwargs.get("sources_filepath", getattr(self, "sources_filepath"))

    _ref_xml = self.XMLBuilder.tostring(_sources_xml, pretty_print=True, encoding='UTF-8', xml_declaration=True)
    self.utility.write(f"{_sources_filepath}/{_sources_filename}", _ref_xml, mode="wb")
    # with open(f"{_sources_filepath}/{_sources_filename}", "wb") as _fh:
    #   _fh.write(_ref_xml)

  def pmid_to_bibitext(self, *args, **kwargs):
    _pmid = args[0] if len(args) > 0 else kwargs.get("pmid", "")
    _destination = args[1] if len(args) > 1 else kwargs.get("destination", "")

    if not len(_pmid) > 4:
      return None

    _headers = {
      "accept": "*/*",
      "accept-encoding": "gzip, deflate, br",
      "accept-language": "en-US,en;q=0.9,hi;q=0.8",
      "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36"
    }
    _payload = {"fromIds": True,"input": f"{_pmid}", "targetFormat":"Bibtex"}
    _res = self.utility.get_file(self.endpoint_pmid_conversion, None, method="post", headers = _headers, json=_payload)
    _json = JSON.loads(_res)
    _ref = REGEX.sub(r'\n\s{10,}', '', _json.get("output"))
    self.utility.write(_destination, _ref)

  def doi_to_bibitext(self, *args, **kwargs):
    self.utility.time_sleep(1)
    _doi = args[0] if len(args) > 0 else kwargs.get("doi", "")
    _destination = args[1] if len(args) > 1 else kwargs.get("destination", "")

    if not "/" in _doi or len(_doi) < 5:
      return None

    if not self.utility.check_path(_destination):
      _url = f"{self.endpoint_doi_conversion}/{_doi}"
      _bibtext = self.utility.get_file(_url, None, headers = {'Accept': 'application/x-bibtex'})
      self.utility.write(_destination, _bibtext)
      return _bibtext
    # @ToDo: If not destination, return the text

    self.utility.log_warning(f"'{_destination}' already exists.")

    return None

  def sanitize_bibitex_text(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text", "")

    _replacements = {
      '--': "-",
    }

    _strips = ["{", "}", ","]
    _text = _text.strip()
    _text = LatexNodes2Text().latex_to_text(_text)
    for _s in _strips:
      _text = _text.strip(_s)

    for _k, _v in _replacements.items():
      _text = _text.replace(_k, _v)

    return _text

  def bibtex_parse_author_names(self, *args, **kwargs):
    _authors = args[0] if len(args) > 0 else kwargs.get("authors", "")
    _authors = _authors.split(" and ")
    _authors_list = []
    for _author in _authors:
      _s = _author.rsplit(" ", 1)
      _fn = _s[0] if len(_s) > 1 else ""
      _ln = _s[1] if len(_s) > 1 else _s[0]
      _authors_list.append((_ln, _fn))
    return _authors_list

  def bibtext_to_dict(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text", "")
    if isinstance(_text, list):
      _text = "".join(_text)

    # Check if multiple @article are present
    _ref_type = ["@incollection", "@article", "@inproceedings"]
    if len(_text) > 10 and any({_m in _text.lower() for _m in _ref_type}):
      _ref_vals = self.regex_parse_bibtext.findall(_text)

      _ref_vals = {_k.strip().lower(): self.sanitize_bibitex_text(_v) for _k, _v in _ref_vals}
      return _ref_vals

    return {}

  def get_reference_details(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text", "")
    _text = _text.replace("https://doi.org/", "DOI:")
    _type = None
    _id = None
    _text = _text.strip()
    _ref_vals = {}

    if _text.lower().startswith("doi"):
      _id = _text[3:] #Split at doi or split at : or ???
      _id = REGEX.sub(r"^\W+", '', _id)
      _id = REGEX.sub(r"\W+$", '', _id)
      _type = "doi"

    elif _text.lower().startswith("pmid"):
      _id = _text[4:]
      _id = REGEX.sub(r"^\W+", '', _id)
      _id = REGEX.sub(r"\W+$", '', _id)
      _type = "pmid"

    elif _text.lower().startswith("pmc"):
      _id = _text[3:]
      _id = REGEX.sub(r"^\W+", '', _id)
      _id = REGEX.sub(r"\W+$", '', _id)
      _type = "pmc"

    if _type is None:
      return None

    # Return placeholder ID to replace with placeholder id
    _placeholder = "".join([_c for _c in _id if _c.isalnum()])
    _placeholder = _placeholder.strip().lower()
    _placeholder = f"{_type}{_placeholder}"

    # _placeholder, dir_references
    # If dir_placeholder holds the bibtext
    _ref_file = f"{self.dir_references}/{_placeholder}.bib"
    if not self.utility.check_path(_ref_file):
      if _type == "doi":
        self.doi_to_bibitext(_id, _ref_file)
      elif _type == "pmid":
        self.pmid_to_bibitext(_id, _ref_file)

    if self.utility.check_path(_ref_file):
      _ref_vals = self.bibtext_to_dict(self.utility.read_text(_ref_file))
      _ref_vals.update({
        "type": _type,
        "identifier": _id.strip(),
        "placeholder": _placeholder,
      })
    else:
      self.utility.log_warning(f"Reference file '{_ref_file}' is missing. This could be latest article which might have not been listed on DOI so need to download reference manually.")

    return _ref_vals

  def process_text(self, *args, **kwargs):
    # @Deprecated
    return
    _text = args[0] if len(args) > 0 else kwargs.get("text", "")
    _ref = []
    _ref1 = REGEX.findall(r'\[(.*?)\]', _text)
    _ref2 = REGEX.findall(r'\((.*?)\)', _text)

    _ref.extend(_ref1)
    _ref.extend(_ref2)

    _ref_det = []

    for _r in _ref:
      _rfd = self.get_reference_details(_r)
      if isinstance(_rfd, dict):
        _rfd and _ref_det.append(_rfd)

        _r_type = _rfd.get("type")
        _r_idf = _rfd.get("identifier")
        _r_idf = "".join([_c for _c in _r_idf if _c.isalnum()])
        _rep = f"{_r_type}{_r_idf}"
        _text = _text.replace(_r, _rep)

    self.references.extend(_ref_det)
    return _text

  def process_paragraphs(self, *args, **kwargs):
    # @Deprecated
    raise Exception("@Deprecated method.")
    _paragraphs = args[0] if len(args) > 0 else kwargs.get("paragraphs", [])
    for _p in _paragraphs:
      _p.text = self.process_text(_p.text)

  def process_tables(self, *args, **kwargs):
    # @Deprecated
    raise Exception("@Deprecated method.")
    _tables = args[0] if len(args) > 0 else kwargs.get("tables", [])
    for _table in _tables:
      for _row in _table.rows:
        for _cell in _row.cells:
          for _p in _cell.paragraphs:
            _p.text = self.process_text(_p.text)

  def process_document_text(self, *args, **kwargs):
    # @Deprecated
    raise Exception("@Deprecated method.")
    _docx = args[0] if len(args) > 0 else kwargs.get("docx")
    if _docx:
      _document = self.Document(_docx)
      self.process_tables(_document.tables)
      self.process_paragraphs(_document.paragraphs)
      _document.save(f"{_docx}-processed.docx")
      return True
    else:
      return None

  def insert_citations(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text")

    self.ref_discovered = REGEX.findall(self.regex_find_ref, _text)
    _ref_cleaned = []

    for _r in self.ref_discovered:
      _r_untagged = REGEX.sub(self.regex_remove_tags, '', _r)
      if len(REGEX.findall(self.regex_validate_ref_text, _r_untagged, REGEX.I)):
        # Replace untagged text to clean broken references
        _text = _text.replace(_r, _r_untagged.strip())
        _ref_cleaned.extend(_r_untagged.strip().split(","))

    for _id, _ref in enumerate(_ref_cleaned):
      _ref_details = self.get_reference_details(_ref.strip())
      if _ref_details is not None and _ref_details.get("title"):

        _random_num = RandInt(10**6, (10**8)-1)

        if not hasattr(self, "citation_hex"):
          _citation_hex = hex(-_random_num).split('x')[-1]
          _citation_hex = _citation_hex[:8].upper()
          _citation_hex = _citation_hex.rjust(8, "0")
          self.citation_hex = _citation_hex

        _random_num = f"-{_random_num}"
        _placeholder = _ref_details.get("placeholder")
        _type = _ref_details.get("type")
        _mod_ref = _placeholder

        # Replace with citation XML String as it requires right markup to replace the citation
        # _mod_ref = self.citation_xml_string.format(citation_placeholder = _placeholder, citation_id = _random_num, citation_hex = self.citation_hex)
        _ref_details.update({
          "citation_hex": self.citation_hex,
        })

        _text = _text.replace(_ref.strip(), _mod_ref)
        self.ref_mapping.append(_ref_details)

    return _text

  def insert_rsids(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text")
    _rsids = args[1] if len(args) > 1 else kwargs.get("rsids")
    for _rsid in _rsids:
      if _rsid not in _text:
        _text = _text.replace("</w:rsids>", f"<w:rsid w:val=\"{_rsid}\"/></w:rsids>")
    return _text

  def insert_placeholders(self, *args, **kwargs):
    _text = args[0] if len(args) > 0 else kwargs.get("text")
    _placeholders = args[1] if len(args) > 1 else kwargs.get("placeholders")
    _placeholder_string = """<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Tag>{placeholder}</b:Tag><b:RefOrder>{ref_order}</b:RefOrder></b:Source>"""

    if not "<b:Sources" in _text:
      return _text

    for _idx, _ph in enumerate(_placeholders):
      if _ph not in _text:
        _text = _text.replace("</b:Sources>", _placeholder_string.format(**{"placeholder": _ph, "ref_order": _idx + 1}))
    return _text

  def process_document_xml(self, *args, **kwargs):
    _docx = args[0] if len(args) > 0 else kwargs.get("docx")

    # WINDOWS_LINE_ENDING = b'\r\n'
    # UNIX_LINE_ENDING = b'\n'

    if _docx:
      _zip_in = self.ZipFile(_docx, 'r')
      _zip_out = self.ZipFile(f"{_docx}-processed.docx", 'w')

      for _zipped_item in _zip_in.infolist():
          _zip_buffer = _zip_in.read(_zipped_item.filename)
          if (_zipped_item.filename == 'word/document.xml'):
            _xml_string = _zip_buffer.decode("utf-8")
            _xml_string = self.insert_citations(_xml_string)
            _zip_buffer = _xml_string.encode("utf-8")
          _zip_out.writestr(_zipped_item, _zip_buffer)

      _zip_out.close()
      _zip_in.close()

    self.ref_mapping = PD.DataFrame(self.ref_mapping)

    """Automatically add placeholder and add source to the document zip. (Currently not working.)"""
    # if self.ref_mapping.shape[0] > 0:
    #   _zip_in = self.ZipFile(f"{_docx}-processed.docx", 'r')
    #   _zip_out = self.ZipFile(f"{_docx}-processed-ref.docx", 'w')

    #   for _zipped_item in _zip_in.infolist():
    #       _zip_buffer = _zip_in.read(_zipped_item.filename)
    #       if (_zipped_item.filename == 'word/settings.xml'):
    #         _xml_string = _zip_buffer.decode("utf-8")
    #         _xml_string = self.insert_rsids(_xml_string, self.ref_mapping["citation_hex"].tolist())
    #         _zip_buffer = _xml_string.encode("utf-8")

    #       if (_zipped_item.filename == 'customXml/item2.xml') or (_zipped_item.filename == 'customXml/item1.xml'):
    #         _xml_string = _zip_buffer.decode("utf-8")
    #         _xml_string = self.insert_placeholders(_xml_string, self.ref_mapping["placeholder"].tolist())
    #         _zip_buffer = _xml_string.encode("utf-8")

    #       _zip_out.writestr(_zipped_item, _zip_buffer)

    #   _zip_out.close()
    #   _zip_in.close()

    self.ref_mapping.to_csv(f"{_docx}-references.csv", index=False)

  def source_add_articles(self, *args, **kwargs):

    _ref_map = self.ref_mapping if isinstance(self.ref_mapping, PD.DataFrame) else PD.DataFrame(self.ref_mapping)
    if "identifier" in _ref_map.columns:
      _ref_map.drop_duplicates(subset=["identifier"], keep="first", inplace=True)
    if "doi" in _ref_map.columns:
      _ref_map.drop_duplicates(subset=["doi"], keep="first", inplace=True)

    for _idx, _ref_row in self.ref_mapping.iterrows():
      _bfp = f"{self.dir_references}/{_ref_row['placeholder']}.bib"
      _tag = self.utility.filename(_bfp)
      _ref_vals = self.bibtext_to_dict(self.utility.read_text(_bfp))

      if not _ref_vals:
        continue

      _idfier = _ref_map[_ref_map["placeholder"] == _tag]
      _idfier = _idfier.iloc[0]["identifier"] if _idfier.shape[0] > 0 else ""

      _ref_vals["tag"] = _tag
      _ref_vals["doi"] = _idfier

      self.sources_add_article({
        "Tag": _ref_vals.get('tag'),
        "SourceType": "JournalArticle",
        "Title": _ref_vals.get('title', ""),
        "Author": self.bibtex_parse_author_names(_ref_vals.get("author", "")),
        "JournalName": _ref_vals.get('journal'),
        "Pages": _ref_vals.get('pages'),
        "Issue": _ref_vals.get('number'),
        "Volume": _ref_vals.get('volume'),
        "Publisher": _ref_vals.get('publisher'),
        "Month": _ref_vals.get('month', "").title(),
        "Year": _ref_vals.get('year'),
        # Day
        # City
        "DOI": _ref_vals.get('doi'),
        "URL": _ref_vals.get('url'),
      })

  # Earlier `def extract()`
  def format_references(self, *args, **kwargs):
    _docx = args[0] if len(args) > 0 else kwargs.get("docx", [])
    _docx = [_docx] if isinstance(_docx, str) else _docx
    for _d in _docx:
      self.process_document_xml(_d)

    self.source_add_articles()
    self.sources_write_xml()

