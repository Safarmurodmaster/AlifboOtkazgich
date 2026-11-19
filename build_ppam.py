"""
build_ppam.py - PowerPoint PPAM faylini Python bilan yaratish
COM automation kerak emas — to'g'ridan-to'g'ri ZIP/OpenXML formatida yasaymiz.

Usul: mavjud AlifboOtkazgich.dotm ni asos sifatib olamiz (u ham ZIP),
undan vbaProject.bin ni ajratib olamiz, keyin ppam strukturasiga solamiz.

Muqobil: xlam fayldan ham qila olamiz — ular bir xil OpenXML formatida.
"""

import os
import sys
import shutil
import zipfile
import struct

def log(msg):
    import datetime
    logdir = os.path.join(os.environ.get('APPDATA',''), 'Alifbo')
    os.makedirs(logdir, exist_ok=True)
    logpath = os.path.join(logdir, 'kurulish-log.txt')
    line = f"{datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}  [PPT-PY] {msg}"
    print(line, flush=True)
    with open(logpath, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    addin_dir = os.path.join(os.environ.get('APPDATA',''), 'Microsoft', 'AddIns')
    os.makedirs(addin_dir, exist_ok=True)
    target = os.path.join(addin_dir, 'AlifboOtkazgich.ppam')

    # Manba sifatida dotm (u ham ZIP/OpenXML)
    word_startup = os.path.join(os.environ.get('APPDATA',''), 'Microsoft', 'Word', 'STARTUP')
    src_dotm = os.path.join(word_startup, 'AlifboOtkazgich.dotm')
    src_xlam = os.path.join(os.environ.get('APPDATA',''), 'Microsoft', 'Excel', 'XLSTART', 'AlifboOtkazgich.xlam')

    # Manba tanlash: dotm yaxshiroq (unda barcha VBA modullar bor)
    src = None
    if os.path.exists(src_dotm) and os.path.getsize(src_dotm) > 10000:
        src = src_dotm
        log(f"Manba: dotm ({os.path.getsize(src_dotm)} bytes)")
    elif os.path.exists(src_xlam) and os.path.getsize(src_xlam) > 10000:
        src = src_xlam
        log(f"Manba: xlam ({os.path.getsize(src_xlam)} bytes)")
    else:
        log("XATO: na dotm, na xlam topilmadi!")
        return 1

    # ppam = pptm bilan bir xil formatda, faqat Content_Types va relationships farq qiladi
    # Strategiya: src ni nusxa olamiz, Content_Types.xml ni almashtitramiz
    tmp = target + '.tmp.zip'
    shutil.copy2(src, tmp)

    ppam_content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/package/2006/metadata/core-properties+xml"/>
  <Override PartName="/ppt/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>'''

    # ZIP ni o'qib, Content_Types.xml ni almashtiramiz
    # Lekin ppam strukturasi pptx bilan farq qiladi — src dotm/xlam
    # Sodda yondashuv: src faylni to'g'ridan-to'g'ri ppam deb nusxa olamiz
    # PowerPoint ppam = pptm (macro-enabled presentation) bilan bir xil format
    # Content_Types.xml dagi main override ContentType ni almashtirish kerak

    try:
        # Mavjud ZIP ni o'qiymiz
        with zipfile.ZipFile(src, 'r') as zin:
            names = zin.namelist()
            log(f"Manba ichidagi fayllar: {len(names)} ta")

            # vbaProject.bin bormi?
            has_vba = any('vbaProject.bin' in n for n in names)
            log(f"vbaProject.bin: {'bor' if has_vba else 'YOQ'}")

        if not has_vba:
            log("XATO: vbaProject.bin topilmadi!")
            return 1

        # Minimal ppam yaratamiz: vbaProject.bin ni dotm dan olib, ppam strukturasiga solamiz
        result = build_ppam_from_vba(src, target, script_dir)
        return result

    except Exception as e:
        log(f"ZIP xato: {e}")
        import traceback
        log(traceback.format_exc())
        return 1
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except: pass


def build_ppam_from_vba(src_path, target_path, script_dir):
    """vbaProject.bin ni src dan olib, to'liq ppam yaratamiz"""
    import io

    # vbaProject.bin ni src dan ajratib olamiz
    vba_data = None
    with zipfile.ZipFile(src_path, 'r') as zin:
        for name in zin.namelist():
            if name.endswith('vbaProject.bin'):
                vba_data = zin.read(name)
                log(f"vbaProject.bin o'lchamlari: {len(vba_data)} bytes")
                break

    if not vba_data:
        log("XATO: vbaProject.bin o'qilmadi")
        return 1

    # Minimal ppam strukturasini yaratamiz
    # ppam = pptm (macro-enabled addin) formatida
    presentation_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''

    presentation_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>'''

    slide_master_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr clr="bg1"/></p:bgRef></p:bg>
  <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>'''

    slide_master_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''

    slide_layout_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>'''

    slide_layout_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

    theme_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
      <a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    </a:fillStyleLst>
    <a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    </a:lnStyleLst>
    <a:effectStyleLst>
      <a:effectStyle><a:effectLst/></a:effectStyle>
      <a:effectStyle><a:effectLst/></a:effectStyle>
      <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
    </a:effectStyleLst>
    <a:bgFillStyleLst>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office PowerPoint</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>'''

    core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Fozliddinov Safarmurod</dc:creator>
  <dc:title>AlifboOtkazgich</dc:title>
</cp:coreProperties>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/package/2006/metadata/core-properties+xml"/>
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    # ZIP fayl yaratamiz
    if os.path.exists(target_path):
        os.remove(target_path)

    try:
        with zipfile.ZipFile(target_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            zout.writestr('[Content_Types].xml', content_types)
            zout.writestr('_rels/.rels', root_rels)
            zout.writestr('ppt/presentation.xml', presentation_xml)
            zout.writestr('ppt/_rels/presentation.xml.rels', presentation_rels)
            zout.writestr('ppt/slideMasters/slideMaster1.xml', slide_master_xml)
            zout.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', slide_master_rels)
            zout.writestr('ppt/slideLayouts/slideLayout1.xml', slide_layout_xml)
            zout.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', slide_layout_rels)
            zout.writestr('ppt/theme/theme1.xml', theme_xml)
            zout.writestr('docProps/app.xml', app_xml)
            zout.writestr('docProps/core.xml', core_xml)
            # VBA binary
            zout.writestr('ppt/vbaProject.bin', vba_data)

        sz = os.path.getsize(target_path)
        log(f"ppam yaratildi: {sz} bytes")
        return 0

    except Exception as e:
        log(f"ZIP yozish xato: {e}")
        import traceback
        log(traceback.format_exc())
        return 1


if __name__ == '__main__':
    sys.exit(main())
