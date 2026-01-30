# -*- coding: utf-8 -*-
import os
import json
import re
import uuid
import threading
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from html import escape
from xml.sax.saxutils import escape as xml_escape
import urllib.request
import urllib.error

import wx
import wx.html

import addonHandler
import api
import globalPluginHandler
import gui
import scriptHandler
import ui
import globalVars
import textInfos
import tones
from logHandler import log

addonHandler.initTranslation()

ADDON_NAME = "readLater"
DATA_DIR_NAME = "readLater"
INDEX_FILE = "index.json"
SETTINGS_FILE = "settings.json"
ARTICLES_DIR = "articles"

DEFAULT_SETTINGS = {
    "preserveFormatting": True,
}

ALLOWED_TAGS = {
    "h1", "h2", "h3", "h4", "h5", "h6",
    "p", "ul", "ol", "li",
    "a", "strong", "em", "b", "i",
    "code", "pre", "blockquote",
    "br",
}
BLOCK_TAGS = {"p", "ul", "ol", "li", "blockquote", "pre", "h1", "h2", "h3", "h4", "h5", "h6"}
SKIP_TAGS = {"script", "style", "noscript", "svg", "canvas", "iframe"}
UNWANTED_TAGS = {"nav", "header", "footer", "aside"}
UNWANTED_CLASS_ID_TOKENS = {
    "nav", "menu", "breadcrumb", "header", "footer", "top", "share", "social",
    "toolbar", "subscribe", "newsletter", "comment", "comments", "advert", "ads",
}
TAG_MAP = {
    "div": "p",
    "section": "p",
    "article": "p",
    "main": "p",
    "header": "p",
    "footer": "p",
    "nav": "p",
    "aside": "p",
}


def _get_data_dir():
    return os.path.join(globalVars.appArgs.configPath, DATA_DIR_NAME)


def _ensure_dirs():
    data_dir = _get_data_dir()
    articles_dir = os.path.join(data_dir, ARTICLES_DIR)
    os.makedirs(articles_dir, exist_ok=True)
    return data_dir, articles_dir


def _load_json(path, default):
    if not os.path.isfile(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _load_index():
    data_dir, _articles_dir = _ensure_dirs()
    return _load_json(os.path.join(data_dir, INDEX_FILE), [])


def _save_index(records):
    data_dir, _articles_dir = _ensure_dirs()
    _save_json(os.path.join(data_dir, INDEX_FILE), records)


def _load_settings():
    data_dir, _articles_dir = _ensure_dirs()
    settings = _load_json(os.path.join(data_dir, SETTINGS_FILE), DEFAULT_SETTINGS)
    merged = DEFAULT_SETTINGS.copy()
    merged.update(settings or {})
    return merged


def _save_settings(settings):
    data_dir, _articles_dir = _ensure_dirs()
    _save_json(os.path.join(data_dir, SETTINGS_FILE), settings)


def _play_save_tone():
    try:
        tones.beep(880, 80)
    except Exception:
        try:
            wx.Bell()
        except Exception:
            pass

def _play_error_tone():
    try:
        tones.beep(220, 200)
    except Exception:
        try:
            wx.Bell()
        except Exception:
            pass


def _maximize_message_window(title):
    try:
        for win in wx.GetTopLevelWindows():
            if win.GetTitle() == title:
                win.Maximize()
                win.Raise()
                win.SetFocus()
                break
    except Exception:
        pass
def _bind_escape_close(dlg):
    def on_char(event):
        if event.GetKeyCode() == wx.WXK_ESCAPE:
            if dlg.IsModal():
                dlg.EndModal(wx.ID_CANCEL)
            else:
                dlg.Close()
            return
        event.Skip()

    dlg.Bind(wx.EVT_CHAR_HOOK, on_char)


class _SimpleHTMLCleaner(HTMLParser):
    def __init__(self):
        super().__init__()
        self.out = []
        self.skip_depth = 0
        self.tag_stack = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        mapped = TAG_MAP.get(tag, tag)
        if tag in SKIP_TAGS:
            self.skip_depth += 1
            return
        if self.skip_depth > 0:
            return
        if mapped not in ALLOWED_TAGS:
            return
        if mapped == "a":
            self.out.append("<a href=\"#\" role=\"link\" aria-disabled=\"true\" class=\"rl-link\">")
            self.tag_stack.append(mapped)
            return
        if mapped == "br":
            self.out.append("<br />")
            return
        self.out.append("<%s>" % mapped)
        self.tag_stack.append(mapped)

    def handle_endtag(self, tag):
        tag = tag.lower()
        mapped = TAG_MAP.get(tag, tag)
        if tag in SKIP_TAGS:
            if self.skip_depth > 0:
                self.skip_depth -= 1
            return
        if self.skip_depth > 0:
            return
        if mapped not in ALLOWED_TAGS or mapped == "br":
            return
        if mapped in self.tag_stack:
            while self.tag_stack:
                open_tag = self.tag_stack.pop()
                self.out.append("</%s>" % open_tag)
                if open_tag == mapped:
                    break

    def handle_data(self, data):
        if self.skip_depth > 0:
            return
        text = data.strip()
        if not text:
            return
        self.out.append(escape(text))

    def get_html(self):
        while self.tag_stack:
            self.out.append("</%s>" % self.tag_stack.pop())
        return "".join(self.out)


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []
        self._last_was_block = False

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in BLOCK_TAGS:
            if self.parts and not self._last_was_block:
                self.parts.append("\n")
            self._last_was_block = True
        if tag == "br":
            self.parts.append("\n")
            self._last_was_block = True

    def handle_data(self, data):
        text = data.strip()
        if not text:
            return
        if self.parts and not self.parts[-1].endswith("\n"):
            self.parts.append(" ")
        self.parts.append(text)
        self._last_was_block = False

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in BLOCK_TAGS:
            self.parts.append("\n")
            self._last_was_block = True

    def get_text(self):
        text = "".join(self.parts)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def _extract_main_html(html):
    # Try to grab <article> or <main> first.
    for tag in ("article", "main"):
        match = re.search(r"<%s\b[^>]*>(.*?)</%s>" % (tag, tag), html, re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1)
    # Fallback to <body>
    match = re.search(r"<body\b[^>]*>(.*?)</body>", html, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1)
    return html


def _clean_html(html, title, url):
    main_html = _extract_main_html(html)
    main_html = _sanitize_html_for_reading(main_html)
    cleaner = _SimpleHTMLCleaner()
    cleaner.feed(main_html)
    cleaned = cleaner.get_html()
    return _inject_header_and_source(cleaned, title, url)


def _html_to_text(clean_html):
    extractor = _TextExtractor()
    extractor.feed(clean_html)
    return extractor.get_text()


def _sanitize_html_for_reading(html):
    # Strip active content and make links non-clickable.
    html = re.sub(r"<(script|style|noscript|iframe|svg|canvas)\b[^>]*>.*?</\1>", "", html, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r"\son\w+\s*=\s*([\"']).*?\1", "", html, flags=re.IGNORECASE | re.DOTALL)
    return html


def _inject_header_and_source(html, title, url):
    has_heading = re.search(r"<h[1-6]\b", html, flags=re.IGNORECASE) is not None
    if not title:
        title = _infer_title_from_html(html) or url or "Article"
    header = "<h1>%s</h1>" % escape(title)
    meta = "<p><strong>Source:</strong> %s</p>" % escape(url) if url else ""
    style = "<style>.rl-link{color:#0066cc;text-decoration:underline;cursor:default;}</style>"
    insert = style + (header if not has_heading else "") + meta
    if not insert:
        return html
    head_match = re.search(r"<head\b[^>]*>", html, flags=re.IGNORECASE)
    if head_match:
        head_idx = head_match.end()
        html = html[:head_idx] + style + html[head_idx:]
        insert = (header if not has_heading else "") + meta
    body_match = re.search(r"<body\b[^>]*>", html, flags=re.IGNORECASE)
    if body_match:
        body_idx = body_match.end()
        return html[:body_idx] + insert + html[body_idx:]
    return "<html><head><meta charset=\"utf-8\"/>%s</head><body>%s%s</body></html>" % (style, insert, html)


def _word_count(text):
    return len(re.findall(r"\b\w+\b", text))


def _fetch_html(url):
    headers = {"User-Agent": "NVDA-Read-Later/0.1"}
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=20) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        return resp.read().decode(charset, errors="replace")


def _write_docx(text, dest_path, title=""):
    paragraphs = [line.strip() for line in text.splitlines() if line.strip()]
    if title:
        paragraphs.insert(0, title)
    document_body = []
    for para in paragraphs:
        document_body.append("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % xml_escape(para))
    document_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        "<w:body>%s</w:body></w:document>" % "".join(document_body)
    )
    content_types = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "</Types>"
    )
    rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
        "Target=\"word/document.xml\"/>"
        "</Relationships>"
    )
    document_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"
    )
    with zipfile.ZipFile(dest_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/_rels/document.xml.rels", document_rels)


def _write_epub(html_body, dest_path, title=""):
    if not title:
        title = "Article"
    chapter = (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        "<html xmlns=\"http://www.w3.org/1999/xhtml\">"
        "<head><title>%s</title></head>"
        "<body>%s</body></html>" % (escape(title), html_body)
    )
    content_opf = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<package xmlns=\"http://www.idpf.org/2007/opf\" unique-identifier=\"BookId\" version=\"2.0\">"
        "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">"
        "<dc:title>%s</dc:title>"
        "<dc:language>en</dc:language>"
        "<dc:identifier id=\"BookId\">urn:uuid:%s</dc:identifier>"
        "</metadata>"
        "<manifest>"
        "<item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>"
        "</manifest>"
        "<spine toc=\"ncx\">"
        "<itemref idref=\"chapter\"/>"
        "</spine>"
        "</package>" % (escape(title), uuid.uuid4())
    )
    container_xml = (
        "<?xml version=\"1.0\"?>"
        "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">"
        "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" 
        "</rootfiles></container>"
    )
    with zipfile.ZipFile(dest_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # epub requires uncompressed mimetype first
        zf.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("META-INF/container.xml", container_xml)
        zf.writestr("OEBPS/content.opf", content_opf)
        zf.writestr("OEBPS/chapter.xhtml", chapter)


class SaveArticleDialog(wx.Dialog):
    def __init__(self, parent, url="", title=""):
        super().__init__(parent, title=_("Save Article"))
        self.url = url
        self.title_value = title
        self.settings = _load_settings()
        _bind_escape_close(self)

        sizer = wx.BoxSizer(wx.VERTICAL)

        title_label = wx.StaticText(self, label=_("Title"))
        self.title_ctrl = wx.TextCtrl(self, value=self.title_value)
        url_label = wx.StaticText(self, label=_("URL"))
        self.url_ctrl = wx.TextCtrl(self, value=self.url)

        self.preserve_check = wx.CheckBox(self, label=_("Preserve formatting (HTML)"))
        self.preserve_check.SetValue(self.settings.get("preserveFormatting", True))

        sizer.Add(title_label, 0, wx.ALL, 5)
        sizer.Add(self.title_ctrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        sizer.Add(url_label, 0, wx.ALL, 5)
        sizer.Add(self.url_ctrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        sizer.Add(self.preserve_check, 0, wx.ALL, 5)

        btn_sizer = self.CreateButtonSizer(wx.OK | wx.CANCEL)
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 10)

        self.SetSizerAndFit(sizer)

    def get_values(self):
        return self.title_ctrl.GetValue().strip(), self.url_ctrl.GetValue().strip(), self.preserve_check.GetValue()


class LibraryDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=_("Read Later Library"), size=(750, 500))
        self.records = _load_index()
        self.filtered = list(self.records)
        _bind_escape_close(self)

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        search_label = wx.StaticText(self, label=_("Search"))
        self.search_ctrl = wx.SearchCtrl(self)
        self.search_ctrl.Bind(wx.EVT_TEXT, self.on_search)

        main_sizer.Add(search_label, 0, wx.ALL, 5)
        main_sizer.Add(self.search_ctrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        self.list_ctrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.list_ctrl.InsertColumn(0, _("Title"), width=260)
        self.list_ctrl.InsertColumn(1, _("Date"), width=120)
        self.list_ctrl.InsertColumn(2, _("Words"), width=80)
        self.list_ctrl.InsertColumn(3, _("URL"), width=260)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_open)

        main_sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.open_btn = wx.Button(self, label=_("Read"))
        self.export_btn = wx.Button(self, label=_("Export"))
        self.export_all_btn = wx.Button(self, label=_("Export All"))
        self.delete_btn = wx.Button(self, label=_("Delete"))
        self.close_btn = wx.Button(self, label=_("Close"))

        self.open_btn.Bind(wx.EVT_BUTTON, self.on_open)
        self.export_btn.Bind(wx.EVT_BUTTON, self.on_export)
        self.export_all_btn.Bind(wx.EVT_BUTTON, self.on_export_all)
        self.delete_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        self.close_btn.Bind(wx.EVT_BUTTON, lambda evt: self.Close())

        for btn in (self.open_btn, self.export_btn, self.export_all_btn, self.delete_btn, self.close_btn):
            button_sizer.Add(btn, 0, wx.ALL, 5)

        main_sizer.Add(button_sizer, 0, wx.ALIGN_CENTER)

        self.SetSizer(main_sizer)
        self._refresh_list()
        wx.CallAfter(self._focus_list)

    def _focus_list(self):
        if self.list_ctrl.GetItemCount() > 0:
            self.list_ctrl.Select(0)
            self.list_ctrl.EnsureVisible(0)
        self.list_ctrl.SetFocus()

    def _refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for record in self.filtered:
            index = self.list_ctrl.InsertItem(self.list_ctrl.GetItemCount(), record.get("title", ""))
            self.list_ctrl.SetItem(index, 1, record.get("dateSaved", ""))
            self.list_ctrl.SetItem(index, 2, str(record.get("wordCount", "")))
            self.list_ctrl.SetItem(index, 3, record.get("url", ""))

    def on_search(self, event):
        term = self.search_ctrl.GetValue().strip().lower()
        if not term:
            self.filtered = list(self.records)
        else:
            self.filtered = [
                r for r in self.records
                if term in (r.get("title", "").lower()) or term in (r.get("url", "").lower())
            ]
        self._refresh_list()

    def _get_selected_record(self):
        idx = self.list_ctrl.GetFirstSelected()
        if idx == -1:
            return None
        return self.filtered[idx]

    def on_open(self, event):
        record = self._get_selected_record()
        if not record:
            ui.message(_("Select an article."))
            return
        data_dir, articles_dir = _ensure_dirs()
        html_path = os.path.join(articles_dir, record["id"] + ".html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()
        except Exception:
            ui.message(_("Unable to open the article file."))
            return
        try:
            title = record.get("title", "Article")
            ui.browseableMessage(html_content, title, True)
            wx.CallAfter(_maximize_message_window, title)
        except Exception:
            ui.message(_("Unable to open the reader view."))

    def _export_record(self, record, dest_path, format_name):
        data_dir, articles_dir = _ensure_dirs()
        html_path = os.path.join(articles_dir, record["id"] + ".html")
        text_path = os.path.join(articles_dir, record["id"] + ".txt")
        with open(html_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        with open(text_path, "r", encoding="utf-8") as f:
            text_content = f.read()
        if format_name == "html":
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(html_content)
        elif format_name == "txt":
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(text_content)
        elif format_name == "md":
            with open(dest_path, "w", encoding="utf-8") as f:
                title = record.get("title", "Article")
                f.write("# %s\n\n%s" % (title, text_content))
        elif format_name == "docx":
            _write_docx(text_content, dest_path, record.get("title", ""))
        elif format_name == "epub":
            # use body without outer HTML
            body_match = re.search(r"<body[^>]*>(.*?)</body>", html_content, re.IGNORECASE | re.DOTALL)
            body = body_match.group(1) if body_match else html_content
            _write_epub(body, dest_path, record.get("title", ""))

    def on_export(self, event):
        record = self._get_selected_record()
        if not record:
            ui.message(_("Select an article."))
            return
        wildcard = "HTML (*.html)|*.html|Text (*.txt)|*.txt|Markdown (*.md)|*.md|DOCX (*.docx)|*.docx|EPUB (*.epub)|*.epub"
        try:
            if gui.mainFrame:
                gui.mainFrame.prePopup()
            with wx.FileDialog(self, message=_("Export Article"), wildcard=wildcard, style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
                if dlg.ShowModal() != wx.ID_OK:
                    return
                path = dlg.GetPath()
                ext = os.path.splitext(path)[1].lower()
                format_map = {".html": "html", ".txt": "txt", ".md": "md", ".docx": "docx", ".epub": "epub"}
                format_name = format_map.get(ext)
                if not format_name:
                    ui.message(_("Unsupported export format."))
                    return
                try:
                    self._export_record(record, path, format_name)
                    ui.message(_("Exported successfully."))
                except Exception:
                    ui.message(_("Export failed."))
        finally:
            if gui.mainFrame:
                gui.mainFrame.postPopup()

    def on_export_all(self, event):
        if not self.records:
            ui.message(_("No articles to export."))
            return
        try:
            if gui.mainFrame:
                gui.mainFrame.prePopup()
            with wx.DirDialog(self, message=_("Export All to Folder")) as dlg:
                if dlg.ShowModal() != wx.ID_OK:
                    return
                folder = dlg.GetPath()
        finally:
            if gui.mainFrame:
                gui.mainFrame.postPopup()
        for record in self.records:
            safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", record.get("title", "article"))
            dest = os.path.join(folder, safe_title + ".html")
            try:
                self._export_record(record, dest, "html")
            except Exception:
                continue
        ui.message(_("Exported all articles to HTML."))

    def on_delete(self, event):
        record = self._get_selected_record()
        if not record:
            ui.message(_("Select an article."))
            return
        if wx.MessageBox(_("Delete selected article?"), _("Confirm"), wx.YES_NO | wx.ICON_QUESTION) != wx.YES:
            return
        data_dir, articles_dir = _ensure_dirs()
        html_path = os.path.join(articles_dir, record["id"] + ".html")
        text_path = os.path.join(articles_dir, record["id"] + ".txt")
        for path in (html_path, text_path):
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except Exception:
                pass
        self.records = [r for r in self.records if r.get("id") != record.get("id")]
        self.filtered = list(self.records)
        _save_index(self.records)
        self._refresh_list()
        ui.message(_("Deleted."))

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Read Later")

    def __init__(self):
        super().__init__()
        self._menu_item = None
        wx.CallAfter(self._add_menu)

    def terminate(self):
        try:
            if self._menu_item:
                gui.mainFrame.sysTrayIcon.toolsMenu.Remove(self._menu_item)
        except Exception:
            pass
        super().terminate()

    def _add_menu(self):
        try:
            menu = gui.mainFrame.sysTrayIcon.toolsMenu
            self._menu_item = menu.Append(wx.ID_ANY, _("Read Later Add-on"))
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self._on_open_library, self._menu_item)
        except Exception:
            self._menu_item = None

    def _pre_popup(self):
        try:
            if gui.mainFrame:
                gui.mainFrame.prePopup()
        except Exception:
            pass

    def _post_popup(self):
        try:
            if gui.mainFrame:
                gui.mainFrame.postPopup()
        except Exception:
            pass

    def _on_open_library(self, event):
        try:
            self._pre_popup()
            dlg = LibraryDialog(gui.mainFrame)
            dlg.ShowModal()
            dlg.Destroy()
        finally:
            self._post_popup()

    @scriptHandler.script(
        description=_("Open the article library"),
        gesture="kb:NVDA+j",
    )
    def script_openLibrary(self, gesture):
        wx.CallAfter(self._on_open_library, None)

    def _get_current_url(self):
        try:
            obj = api.getFocusObject()
            if obj is None:
                return ""
            ti = getattr(obj, "treeInterceptor", None)
            if ti:
                url = getattr(ti, "documentConstantIdentifier", None)
                if url and isinstance(url, str) and url.startswith("http"):
                    return url
            for attr in ("url", "URL", "value"):
                url = getattr(obj, attr, None)
                if url and isinstance(url, str) and url.startswith("http"):
                    return url
        except Exception:
            log.exception("Failed to read current URL")
        return ""

    def _get_current_title(self):
        try:
            obj = api.getFocusObject()
            if obj and getattr(obj, "name", None):
                return obj.name
        except Exception:
            log.exception("Failed to read current title")
        return ""

    def _get_focus_text_snapshot(self):
        obj = api.getFocusObject()
        ti = getattr(obj, "treeInterceptor", None)
        if not ti:
            return ""
        try:
            text_info = ti.makeTextInfo(textInfos.POSITION_ALL)
            return text_info.text
        except Exception:
            return ""

    @scriptHandler.script(
        description=_("Save current article"),
        gesture="kb:NVDA+alt+d",
    )
    def script_saveArticle(self, gesture):
        url = self._get_current_url()
        title = self._get_current_title()
        focus_text = self._get_focus_text_snapshot()
        wx.CallAfter(self._show_save_dialog, url, title, focus_text)

    def _show_save_dialog(self, url, title, focus_text):
        dlg = None
        popup_open = False
        try:
            if not gui.mainFrame:
                ui.message(_("NVDA UI is not available."))
                return
            self._pre_popup()
            popup_open = True
            dlg = SaveArticleDialog(gui.mainFrame, url=url, title=title)
            result = dlg.ShowModal()
            if result == wx.ID_OK:
                title, url, preserve = dlg.get_values()
                settings = _load_settings()
                settings["preserveFormatting"] = preserve
                _save_settings(settings)
                if not url:
                    ui.message(_("URL is required."))
                    return
                wx.CallAfter(ui.message, _("Saving article, please wait..."))
                thread = threading.Thread(target=self._save_article_worker, args=(title, url, preserve, focus_text))
                thread.daemon = True
                thread.start()
        except Exception:
            log.exception("Save Article dialog failed")
            ui.message(_("Save Article failed."))
        finally:
            if popup_open:
                self._post_popup()
            try:
                dlg.Destroy()
            except Exception:
                pass

    def _save_article_worker(self, title, url, preserve, focus_text):
        try:
            html = _fetch_html(url)
            if not title:
                title = _infer_title_from_html(html) or url
            clean_html = _clean_html(html, title, url) if preserve else _make_plain_html(title, url, html)
            text_content = _html_to_text(clean_html)
            word_count = _word_count(text_content)

            data_dir, articles_dir = _ensure_dirs()
            article_id = uuid.uuid4().hex
            html_path = os.path.join(articles_dir, article_id + ".html")
            text_path = os.path.join(articles_dir, article_id + ".txt")

            with open(html_path, "w", encoding="utf-8") as f:
                f.write(clean_html)
            with open(text_path, "w", encoding="utf-8") as f:
                f.write(text_content)

            record = {
                "id": article_id,
                "title": title,
                "url": url,
                "dateSaved": datetime.now().strftime("%Y-%m-%d"),
                "wordCount": word_count,
            }
            records = _load_index()
            records.insert(0, record)
            _save_index(records)

            wx.CallAfter(ui.message, _("Article saved."))
            wx.CallAfter(_play_save_tone)
        except urllib.error.URLError:
            if focus_text:
                try:
                    clean_html = _make_plain_html(title or "Article", url, focus_text)
                    text_content = focus_text.strip()
                    word_count = _word_count(text_content)
                    data_dir, articles_dir = _ensure_dirs()
                    article_id = uuid.uuid4().hex
                    html_path = os.path.join(articles_dir, article_id + ".html")
                    text_path = os.path.join(articles_dir, article_id + ".txt")
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(clean_html)
                    with open(text_path, "w", encoding="utf-8") as f:
                        f.write(text_content)
                    record = {
                        "id": article_id,
                        "title": title or "Article",
                        "url": url,
                        "dateSaved": datetime.now().strftime("%Y-%m-%d"),
                        "wordCount": word_count,
                    }
                    records = _load_index()
                    records.insert(0, record)
                    _save_index(records)
                    wx.CallAfter(ui.message, _("Article saved using on-screen text."))
                    wx.CallAfter(_play_save_tone)
                    return
                except Exception:
                    pass
            wx.CallAfter(ui.message, _("Failed to download the page."))
            wx.CallAfter(_play_error_tone)
        except Exception:
            wx.CallAfter(ui.message, _("Saving failed."))
            wx.CallAfter(_play_error_tone)


def _infer_title_from_html(html):
    match = re.search(r"<title\b[^>]*>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    if match:
        text = re.sub(r"\s+", " ", match.group(1))
        return text.strip()
    return ""


def _make_plain_html(title, url, html):
    # When formatting is not preserved, keep only plain text.
    text = _strip_tags(html)
    body = "<h1>%s</h1>" % escape(title or "Article")
    if url:
        body += "<p><strong>Source:</strong> %s</p>" % escape(url)
    body += "<pre>%s</pre>" % escape(text)
    return "<html><head><meta charset=\"utf-8\"/></head><body>%s</body></html>" % body


def _strip_tags(html):
    # Tolerate malformed closing tags like </script\t\n foo> by allowing extra text.
    text = re.sub(r"<script\b[^>]*>.*?</script\b[^>]*>", "", html, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<style\b[^>]*>.*?</style\b[^>]*>", "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()
