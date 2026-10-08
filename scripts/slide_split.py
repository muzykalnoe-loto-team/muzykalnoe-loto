# Один слайд старой игры → отдельный .pptx, без PowerPoint: копия архива, в которой оставлен
# только этот слайд и всё, на что он ссылается (звук, картинки, макет, тема)
import posixpath
import re
import zipfile

SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
CONTENT_TYPES = "[Content_Types].xml"


def _rels_path(part: str) -> str:
    base, name = posixpath.split(part)
    return f"{base}/_rels/{name}.rels" if base else f"_rels/{name}.rels"


def _rel_tags(xml: str) -> list[str]:
    return re.findall(r"<Relationship\b[^>]*>", xml)


def _attr(tag: str, name: str) -> str:
    m = re.search(rf'\b{name}="([^"]*)"', tag)
    return m.group(1) if m else ""


def _resolve(base: str, target: str) -> str:
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))


def extract_slide(src: str, slide_part: str, out_path: str) -> None:
    with zipfile.ZipFile(src) as zin:
        names = set(zin.namelist())
        if slide_part not in names:
            raise ValueError(f"В {src} нет слайда {slide_part}")
        changed: dict[str, str] = {}

        # 1. Связи презентации: убрать все слайды, кроме нужного
        pres_rels_name = "ppt/_rels/presentation.xml.rels"
        pres_rels = zin.read(pres_rels_name).decode("utf-8")
        keep_rid = None
        for tag in _rel_tags(pres_rels):
            if _attr(tag, "Type") != SLIDE_REL:
                continue
            if _resolve("ppt", _attr(tag, "Target")) == slide_part:
                keep_rid = _attr(tag, "Id")
            else:
                pres_rels = pres_rels.replace(tag, "")
        changed[pres_rels_name] = pres_rels

        # 2. Список слайдов, разделы и произвольные показы в presentation.xml
        pres = zin.read("ppt/presentation.xml").decode("utf-8")
        keep_id = None
        for tag in re.findall(r"<p:sldId\b[^>]*/>", pres):
            if f'r:id="{keep_rid}"' in tag:
                keep_id = _attr(tag, "id")
            else:
                pres = pres.replace(tag, "")
        pres = re.sub(r'<p14:sldId id="(\d+)"\s*/>', lambda m: m.group(0) if m.group(1) == keep_id else "", pres)
        pres = re.sub(r"<p:custShowLst>.*?</p:custShowLst>", "", pres, flags=re.S)
        changed["ppt/presentation.xml"] = pres

        # 3. Ссылки нашего слайда на другие слайды → на него самого
        slide_rels_name = _rels_path(slide_part)
        if slide_rels_name in names:
            srels = zin.read(slide_rels_name).decode("utf-8")
            own = posixpath.basename(slide_part)
            for tag in _rel_tags(srels):
                if _attr(tag, "Type") == SLIDE_REL:
                    srels = srels.replace(tag, re.sub(r'\bTarget="[^"]*"', f'Target="{own}"', tag))
            changed[slide_rels_name] = srels

        def read_text(name: str) -> str:
            if name in changed:
                return changed[name]
            xml = zin.read(name).decode("utf-8", "replace")
            # Ссылки на другие слайды из макета, шаблона, заметок (кнопка «Таблица» и т. п.) → на наш
            # слайд: иначе по ним в файл попадает вся игра
            owner_base = posixpath.dirname(posixpath.dirname(name))
            own_rel = posixpath.relpath(slide_part, owner_base or ".")
            for tag in _rel_tags(xml):
                if _attr(tag, "Type") == SLIDE_REL and _resolve(owner_base, _attr(tag, "Target")) != slide_part:
                    xml = xml.replace(tag, re.sub(r'\bTarget="[^"]*"', f'Target="{own_rel}"', tag))
            if name != "_rels/.rels":
                changed[name] = xml
            return xml

        # 4. Обход связей от корня: оставляем только то, до чего можно дойти
        keep: set[str] = set()
        queue = ["_rels/.rels"]
        while queue:
            rels_name = queue.pop()
            if rels_name in keep or rels_name not in names:
                continue
            keep.add(rels_name)
            owner_base = posixpath.dirname(posixpath.dirname(rels_name))
            for tag in _rel_tags(read_text(rels_name)):
                if _attr(tag, "TargetMode") == "External":
                    continue
                part = _resolve(owner_base, _attr(tag, "Target"))
                if part in names and part not in keep:
                    keep.add(part)
                    queue.append(_rels_path(part))

        # 5. Типы содержимого — без удалённых частей
        ct = zin.read(CONTENT_TYPES).decode("utf-8")
        ct = re.sub(r'<Override\b[^>]*PartName="/([^"]+)"[^>]*/>',
                    lambda m: m.group(0) if m.group(1) in keep else "", ct)
        changed[CONTENT_TYPES] = ct
        keep.add(CONTENT_TYPES)

        with zipfile.ZipFile(out_path, "w") as zout:
            for info in zin.infolist():
                if info.filename not in keep:
                    continue
                if info.filename in changed:
                    zout.writestr(info, changed[info.filename].encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
                else:
                    with zin.open(info) as fsrc, zout.open(info, "w") as fdst:
                        while chunk := fsrc.read(1 << 20):
                            fdst.write(chunk)
