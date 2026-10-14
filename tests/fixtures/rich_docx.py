"""Сборка эталонного документа DOCX из сырого OOXML.

Документ намеренно содержит элементы, которых нет в API python-docx:
сноски, комментарии, поле автоматического содержания, объект OLE,
элемент управления содержимым (sdt) и пользовательскую XML-часть.
Такой файл нужен, чтобы проверять сохранность содержимого: потеря
любого элемента при обработке становится видимой ошибкой теста.

Пакет собирается вручную (zip), а не средствами python-docx, иначе
проверка сохранности проверяла бы сама себя.
"""

from __future__ import annotations

import base64
import zipfile
from pathlib import Path

# 1x1 PNG. Байты проверены открытием в Pillow перед добавлением сюда.
PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGM4IScHAAK2AQU0pnWqAAAAAElFTkSuQmCC"
)

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"

XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

CONTENT_TYPES = (
    XML_HEADER
    + f"""<Types xmlns="{NS_CT}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>
</Types>"""
)

PACKAGE_RELS = (
    XML_HEADER
    + f"""<Relationships xmlns="{NS_REL}">
<Relationship Id="rId1" Type="{NS_R}/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="{NS_R}/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId4" Type="{NS_R}/customXml" Target="customXml/item1.xml"/>
</Relationships>"""
)

DOCUMENT_RELS = (
    XML_HEADER
    + f"""<Relationships xmlns="{NS_REL}">
<Relationship Id="rId1" Type="{NS_R}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{NS_R}/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="{NS_R}/footnotes" Target="footnotes.xml"/>
<Relationship Id="rId4" Type="{NS_R}/comments" Target="comments.xml"/>
<Relationship Id="rId5" Type="{NS_R}/header" Target="header1.xml"/>
<Relationship Id="rId6" Type="{NS_R}/footer" Target="footer1.xml"/>
<Relationship Id="rId7" Type="{NS_R}/image" Target="media/image1.png"/>
<Relationship Id="rId8" Type="{NS_R}/hyperlink" Target="https://www.fa.ru/" TargetMode="External"/>
<Relationship Id="rId9" Type="{NS_R}/oleObject" Target="embeddings/oleObject1.bin"/>
</Relationships>"""
)

STYLES = (
    XML_HEADER
    + f"""<w:styles xmlns:w="{NS_W}">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="Heading1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Heading1"/><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2E74B5"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="1F4D78"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:link w:val="FootnoteTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar"><w:name w:val="Footnote Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="FootnoteText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>"""
)

SETTINGS = (
    XML_HEADER
    + f"""<w:settings xmlns:w="{NS_W}">
<w:zoom w:percent="100"/>
<w:defaultTabStop w:val="708"/>
<w:characterSpacingControl w:val="doNotCompress"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""
)

FOOTNOTES = (
    XML_HEADER
    + f"""<w:footnotes xmlns:w="{NS_W}" xmlns:r="{NS_R}">
<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Пояснение в подстрочной сноске.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""
)

COMMENTS = (
    XML_HEADER
    + f"""<w:comments xmlns:w="{NS_W}">
<w:comment w:id="1" w:author="Рецензент" w:date="2026-01-01T00:00:00Z" w:initials="Р"><w:p><w:r><w:t>Проверить формулировку абзаца.</w:t></w:r></w:p></w:comment>
</w:comments>"""
)

HEADER = (
    XML_HEADER
    + f"""<w:hdr xmlns:w="{NS_W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Краснодарский филиал Финуниверситета</w:t></w:r></w:p></w:hdr>"""
)

FOOTER = (
    XML_HEADER
    + f"""<w:ftr xmlns:w="{NS_W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>"""
)

SECT_PR_FIRST = f"""<w:sectPr>
<w:headerReference w:type="default" r:id="rId5"/>
<w:footerReference w:type="default" r:id="rId6"/>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="1134" w:footer="709" w:gutter="0"/>
<w:cols w:space="708"/>
<w:docGrid w:linePitch="360"/>
</w:sectPr>"""

SECT_PR_LAST = f"""<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="1134" w:footer="709" w:gutter="0"/>
<w:cols w:space="708"/>
<w:docGrid w:linePitch="360"/>
</w:sectPr>"""

DOCUMENT = (
    XML_HEADER
    + f"""<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}" xmlns:v="{NS_V}" xmlns:o="{NS_O}" xmlns:w10="urn:schemas-microsoft-com:office:word">
<w:body>
<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>КРАСНОДАРСКИЙ ФИЛИАЛ ФИНУНИВЕРСИТЕТА</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Курсовая работа (тестовый документ)</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Оглавление обновится при открытии</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>ВВЕДЕНИЕ</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Обычный текст с ручным шрифтом, </w:t></w:r><w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t>курсивом</w:t></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> и </w:t></w:r><w:r><w:rPr><w:u w:val="single"/><w:sz w:val="24"/></w:rPr><w:t>подчёркиванием</w:t></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 Понятие и сущность процесса</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:firstLine="709"/></w:pPr><w:r><w:t xml:space="preserve">Ссылка на источник: </w:t></w:r><w:hyperlink r:id="rId8"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>сайт Финуниверситета</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> и текст с </w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t>сноской.</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Таблица 1 – Показатели деятельности</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2600"/><w:gridCol w:w="2600"/><w:gridCol w:w="2600"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Показатель</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Значение</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Единица</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Выручка</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>55000</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>руб.</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Прочие данные</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>–</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>х</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="190500" cy="190500"/><wp:docPr id="1" name="Рисунок 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Рисунок 1 – Схема процесса</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440"><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:72pt" o:ole=""><v:imagedata r:id="rId7" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.3" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_123456" r:id="rId9"/></w:object></w:r></w:p>
<w:sdt><w:sdtPr><w:alias w:val="Блок сведений об авторе"/><w:id w:val="1234567"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Содержимое элемента управления</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Абзац, к которому оставлен комментарий.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r></w:p>
<w:p><w:r><w:t>Текст перед разрывом страницы.</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>
<w:p><w:pPr>{SECT_PR_FIRST}</w:pPr></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>ЗАКЛЮЧЕНИЕ</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:firstLine="709"/></w:pPr><w:r><w:t>Текст заключения во втором разделе документа.</w:t></w:r></w:p>
{SECT_PR_LAST}
</w:body>
</w:document>"""
)

CORE_PROPS = (
    XML_HEADER
    + f"""<cp:coreProperties xmlns:cp="{NS_CP}" xmlns:dc="{NS_DC}" xmlns:dcterms="{NS_DCTERMS}" xmlns:xsi="{NS_XSI}">
<dc:title>Курсовая работа (тестовый документ)</dc:title>
<dc:creator>Тестовый автор</dc:creator>
<cp:lastModifiedBy>Тестовый автор</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""
)

APP_PROPS = (
    XML_HEADER
    + """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Word Formatter test fixture</Application>
<Pages>1</Pages>
<Paragraphs>12</Paragraphs>
</Properties>"""
)

CUSTOM_XML = (
    XML_HEADER
    + '<root xmlns="urn:word-formatter:test-fixture"><value>эту часть приложение трогать не должно</value></root>'
)

# Части, которых python-docx не понимает: он хранит их как есть и обязан
# перенести в результат байт в байт. Проверено 2026-10: так и происходит.
BYTE_IDENTICAL_PARTS = (
    "word/footnotes.xml",
    "word/media/image1.png",
    "word/embeddings/oleObject1.bin",
    "customXml/item1.xml",
)

# Части, которые python-docx разбирает и записывает заново: в файле меняются
# отступы и кавычки XML-заголовка, но состав элементов обязан совпадать.
# Именно поэтому их сравнивают по структуре, а не по байтам.
RESERIALIZED_PARTS = (
    "word/document.xml",
    "word/styles.xml",
    "word/settings.xml",
    "word/comments.xml",
    "word/header1.xml",
    "word/footer1.xml",
    "docProps/core.xml",
    "docProps/app.xml",
)

PARTS: dict[str, bytes] = {
    "[Content_Types].xml": CONTENT_TYPES.encode("utf-8"),
    "_rels/.rels": PACKAGE_RELS.encode("utf-8"),
    "word/document.xml": DOCUMENT.encode("utf-8"),
    "word/_rels/document.xml.rels": DOCUMENT_RELS.encode("utf-8"),
    "word/styles.xml": STYLES.encode("utf-8"),
    "word/settings.xml": SETTINGS.encode("utf-8"),
    "word/footnotes.xml": FOOTNOTES.encode("utf-8"),
    "word/comments.xml": COMMENTS.encode("utf-8"),
    "word/header1.xml": HEADER.encode("utf-8"),
    "word/footer1.xml": FOOTER.encode("utf-8"),
    "word/media/image1.png": PNG_1X1,
    "word/embeddings/oleObject1.bin": b"FAKE-OLE-PAYLOAD-\x00\x01\x02",
    "customXml/item1.xml": CUSTOM_XML.encode("utf-8"),
    "docProps/core.xml": CORE_PROPS.encode("utf-8"),
    "docProps/app.xml": APP_PROPS.encode("utf-8"),
}


def build_rich_docx(target: str | Path) -> Path:
    """Собирает эталонный DOCX и возвращает путь к нему."""
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in PARTS.items():
            archive.writestr(name, data)
    return target
