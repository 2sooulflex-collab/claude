"""Убирает из xlsx необязательные части (тему, docProps), чтобы файл для загрузки через base64 был меньше."""
import re
import sys
import zipfile

DROP = ("docProps/app.xml", "docProps/core.xml", "xl/theme/theme1.xml")


def strip(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for info in zin.infolist():
            name = info.filename
            if name in DROP:
                continue
            data = zin.read(name).decode("utf-8")
            if name == "[Content_Types].xml":
                data = re.sub(r'<Override PartName="/(docProps/[^"]+|xl/theme/theme1\.xml)"[^>]*/>', "", data)
            elif name == "_rels/.rels":
                data = re.sub(r'<Relationship [^>]*Target="/?docProps/[^"]+"[^>]*/>', "", data)
            elif name == "xl/_rels/workbook.xml.rels":
                data = re.sub(r'<Relationship [^>]*Target="/?(xl/)?theme/theme1\.xml"[^>]*/>', "", data)
            z.writestr(name, data)


if __name__ == "__main__":
    strip(sys.argv[1], sys.argv[2])
