"""Set Russian proofing language on slide text (speaker notes stay English)."""
import re
import sys
import zipfile
import shutil

src = sys.argv[1]
tmp = src + ".tmp"
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        name = item.filename
        if re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", name):
            data = data.replace(b'lang="en-US"', b'lang="ru-RU"')
        elif re.match(r"ppt/charts/chart\d+\.xml$", name):
            data = data.replace(b'<c:lang val="en-US"/>', b'<c:lang val="ru-RU"/>').replace(b'lang="en-US"', b'lang="ru-RU"')
        zout.writestr(item, data)
shutil.move(tmp, src)
print("lang fixed:", src)
