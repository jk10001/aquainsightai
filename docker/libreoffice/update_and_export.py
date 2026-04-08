# /app/update_and_export.py
import os, sys, uno
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException


SAVE_BACK_TO_DOCX = False # ⚠️ If you set this True, LibreOffice may flatten/remove some Word field codes on DOCX save.
ATTEMPT_FIELD_REFRESH = False #Libreoffice field refresh not functional currently due to Libreoffice limitations in docx feild code support. 


def _prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _to_url(path):
    p = os.path.abspath(path).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p
    return "file://" + p


def _dispatch_update_all(smgr, ctx, doc):
    frame = doc.getCurrentController().getFrame()
    dispatcher = smgr.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx
    )
    for cmd in (".uno:UpdateAllIndexes", ".uno:UpdateAll", ".uno:UpdateFields"):
        try:
            dispatcher.executeDispatch(frame, cmd, "", 0, ())
        except UnoException:
            pass


def _refresh_text_fields(obj):
    try:
        obj.getTextFields().refresh()
    except Exception:
        pass


def _update_document_indexes(doc):
    try:
        idxs = doc.getDocumentIndexes()
        for i in range(idxs.getCount()):
            try:
                idxs.getByIndex(i).update()
            except Exception:
                pass
    except Exception:
        pass


def _refresh_headers_footers(doc):
    try:
        fams = doc.getStyleFamilies()
        page_styles = fams.getByName("PageStyles")
        for name in page_styles.getElementNames():
            style = page_styles.getByName(name)
            try:
                if getattr(style, "HeaderIsOn", False):
                    _refresh_text_fields(style.HeaderText)
            except Exception:
                pass
            try:
                if getattr(style, "FooterIsOn", False):
                    _refresh_text_fields(style.FooterText)
            except Exception:
                pass
    except Exception:
        pass


def _refresh_text_frames(doc):
    try:
        frames = doc.getTextFrames()
        for n in frames.getElementNames():
            try:
                _refresh_text_fields(frames.getByName(n).getText())
            except Exception:
                pass
    except Exception:
        pass


def main(in_docx_path, out_pdf_path):
    if not os.path.isfile(in_docx_path) or os.path.getsize(in_docx_path) == 0:
        print(f"[ERR] Input missing/empty: {in_docx_path}")
        sys.exit(2)

    in_url = _to_url(in_docx_path)
    out_url = _to_url(out_pdf_path)

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve(
        "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"
    )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # Force DOCX import filter to avoid type-detection flakes
    load_props = (
        _prop("Hidden", True),
        _prop("AsTemplate", False),
        _prop("ReadOnly", False),
        _prop("FilterName", "MS Word 2007 XML"),
    )

    doc = desktop.loadComponentFromURL(in_url, "_blank", 0, load_props)
    try:
        if ATTEMPT_FIELD_REFRESH:
            # API-level refresh
            _update_document_indexes(doc)
            _refresh_text_fields(doc)  # main story
            _refresh_headers_footers(doc)  # headers/footers
            _refresh_text_frames(doc)  # text frames

            # Dispatch passes (belt-and-braces)
            _dispatch_update_all(smgr, ctx, doc)
            _dispatch_update_all(smgr, ctx, doc)

            # ⚠️ Avoid saving back to DOCX by default (LO may remove Word field codes)
            if SAVE_BACK_TO_DOCX:
                doc.store()

        # Export to PDF (this captures the refreshed fields)
        pdf_props = (_prop("FilterName", "writer_pdf_Export"), _prop("Overwrite", True))
        doc.storeToURL(out_url, pdf_props)

    finally:
        doc.close(True)


if __name__ == "__main__":
    if len(sys.argv) != 2 and len(sys.argv) != 3:
        print("Usage: /usr/bin/python3 /app/update_and_export.py IN.docx OUT.pdf")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
