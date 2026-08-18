from html import escape
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


OUT = Path("hong_kong_biomedical_engineering_undergraduate_guide.docx")


def xml_text(text):
    return escape(text, quote=False)


def p(text="", style=None, bold=False, italic=False, color=None, size=None, num=None, keep_next=False):
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if num is not None:
        ppr.append(
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%s"/></w:numPr>' % num
        )
    if keep_next:
        ppr.append("<w:keepNext/>")
    rpr = []
    if bold:
        rpr.append("<w:b/>")
    if italic:
        rpr.append("<w:i/>")
    if color:
        rpr.append(f'<w:color w:val="{color}"/>')
    if size:
        rpr.append(f'<w:sz w:val="{int(size * 2)}"/>')
        rpr.append(f'<w:szCs w:val="{int(size * 2)}"/>')
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
    return f"<w:p>{ppr_xml}<w:r>{rpr_xml}<w:t>{xml_text(text)}</w:t></w:r></w:p>"


def cell(text, width, shading=None, bold=False, align="left"):
    shd = f'<w:shd w:fill="{shading}"/>' if shading else ""
    jc = f'<w:jc w:val="{align}"/>' if align != "left" else ""
    parts = []
    for line in text.split("\n"):
        parts.append(p(line, bold=bold, size=9.5))
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
        '<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>'
        '<w:start w:w="140" w:type="dxa"/><w:end w:w="140" w:type="dxa"/></w:tcMar>'
        f"{jc}<w:vAlign w:val=\"center\"/></w:tcPr>{''.join(parts)}</w:tc>"
    )


def table(headers, rows, widths):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl = [
        '<w:tbl><w:tblPr><w:tblStyle w:val="CompactTable"/>'
        '<w:tblW w:w="9360" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="B8C2CC"/>'
        '<w:left w:val="single" w:sz="4" w:color="B8C2CC"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="B8C2CC"/>'
        '<w:right w:val="single" w:sz="4" w:color="B8C2CC"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="D6DDE5"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="D6DDE5"/></w:tblBorders>'
        '</w:tblPr><w:tblGrid>',
        grid,
        "</w:tblGrid>",
    ]
    tbl.append("<w:tr>" + "".join(cell(h, w, "E8EEF5", True) for h, w in zip(headers, widths)) + "</w:tr>")
    for row in rows:
        tbl.append("<w:tr>" + "".join(cell(str(v), w) for v, w in zip(row, widths)) + "</w:tr>")
    tbl.append("</w:tbl>")
    return "".join(tbl)


def section_break():
    return (
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:pPr></w:p>"
    )


top_rows = [
    ["1", "HKU", "University of Hong Kong", "QS WUR 2026: 11"],
    ["2", "CUHK", "Chinese University of Hong Kong", "QS WUR 2026: 32"],
    ["3", "HKUST", "Hong Kong University of Science and Technology", "QS WUR 2026: 44"],
    ["4", "PolyU", "Hong Kong Polytechnic University", "QS WUR 2026: 54"],
    ["5", "CityUHK", "City University of Hong Kong", "QS WUR 2026: 63"],
]

program_rows = [
    ["HKU", "BEng in Biomedical Engineering (JS6925)", "Direct BME programme jointly offered by Engineering and HKUMed.", "HKDSE: English 3, Chinese 3, Math 3, CSD Attained/LS 2, plus two electives at 3; one elective must be Biology, Chemistry, Physics or Combined Science. M1/M2 counts as a full elective and is preferred.", "Non-JUPAS: high-school Mathematics plus Biology, Chemistry or Physics. Competitive reference: IB 36/45 with HL Math 6 and science requirement; GCE/IAL 2A*1A with A in Math/Further/Pure Math and A in Bio/Chem/Physics."],
    ["CUHK", "BEng in Biomedical Engineering (JS4460)", "Engineering programme with Medicine collaboration and focus areas in instrumentation, imaging/informatics/modelling, and molecular/cell/tissue engineering.", "HKDSE: Chinese 3, English 3, Math 3, CSD Attained, two electives at 3; at least one elective should be M1/M2, Biology, Chemistry or Physics. BME score uses Best 5 with weighting for English and science/math subjects.", "Non-JUPAS/international: applicants with GCE, IB, SAT/AP and other qualifications can apply. Preference is given to strong grades in at least two of Physics, Chemistry, Biology or Mathematics. Senior-year route is available for suitable AD/HD holders."],
    ["HKUST", "BEng in Bioengineering via Department of Chemical and Biological Engineering (JS5220); related BSc Biomedical and Health Sciences (JS5118)", "HKUST route is related rather than titled BME. Bioengineering sits in engineering; Biomedical and Health Sciences is a science/health route.", "Engineering HKDSE minimum for JS5220: English 3, Chinese 3, Math 3, CSD Attained, one of Biology/Chemistry/Physics/ICT at 3, plus one Category A/M1/M2 subject at 3. JS5220 score formula emphasizes English, Math and science subjects.", "For JS5118 BSc Biomedical and Health Sciences: HKDSE English 3, Chinese 3, Math 2, CSD Attained, Elective 1 from Bio/Chem/Physics/M1/M2 at 3, Elective 2 from Category A/M1/M2 at 3; compulsory interview noted by the programme."],
    ["PolyU", "BSc (Hons) in Biomedical Engineering (JS3150)", "Accredited BME programme with Common Year One in Engineering; streams include biomedical engineering and prosthetics/orthotics-related study.", "HKDSE: satisfy PolyU general entrance requirements: Chinese 3, English 3, Math 2, CSD Attained, two electives at 3. No compulsory subject requirement for JS3150, but preferred/highest-weighting subjects include Math, M1/M2, Biology, Chemistry, Physics, Combined Science, DAT, ICT and relevant Applied Learning subjects.", "Admission score: Any Best 5 subjects with programme subject weighting. 2025/26 average weighted score reported by PolyU: 226.0. Non-JUPAS and international applicants follow PolyU general, English and qualification-specific requirements."],
    ["CityUHK", "BEng Biomedical Engineering (JS1211)", "Direct four-year BME programme with features in medical technology, bioinstrumentation, cell and tissue engineering, and biomedical robotics.", "HKDSE: Chinese 3, English 3, Math 2, CSD Attained, plus any one subject at 3 and one of Biology, Chemistry or Physics at 3. M1/M2 and Category C other languages may be used for elective requirements; Applied Learning is not counted as an elective.", "Direct/Non-JUPAS: applicants must satisfy general entrance requirements and are expected to have a science or engineering background. Advanced Standing applicants are expected to have AD/HD with CGPA at least 3.0 or equivalent in a science/engineering-related discipline; Physics and Math backgrounds preferred."],
]

checklist = [
    "Build a science-heavy profile. For BME routes, Biology, Chemistry and Physics are the most repeatedly named subjects; Mathematics and M1/M2 are especially useful for engineering-weighted scoring.",
    "Do not rely on minimum levels as target scores. Minimum entrance requirements only establish eligibility; competitive admitted scores are usually higher and can move each year.",
    "Check route-specific rules before applying. JUPAS, Non-JUPAS, international qualifications, Mainland Gaokao and senior-year/sub-degree entry can have different requirements.",
    "Treat HKUST separately. Choose Bioengineering if the target is engineering design/process/biological systems; choose Biomedical and Health Sciences if the target is a science/health pathway.",
    "Look for interviews and programme-specific assessment. HKU, HKUST BMH, CityUHK and other programmes may use interviews, tests or holistic factors where listed.",
    "Re-check official pages before final submission. This guide uses 2026/27 official admissions pages available on 12 Aug 2026, but requirements can change by entry year.",
]

sources = [
    ["Top-five HK universities", "HK Talent Engage summary of QS World University Rankings 2026", "https://www.hkengage.gov.hk/en/media/news/qs-world-universities-rankings-five-hong-kong-universities-in-the-top-one-hundred"],
    ["HKU BEng(BME)", "HKUMed Undergraduate Admissions - BEng(BME), JS6925", "https://hkumed-ugadmissions.hku.hk/ug_programmes/bme/"],
    ["CUHK BME", "CUHK Department of Biomedical Engineering - Bachelor admission", "https://www.bme.cuhk.edu.hk/new/bachelor.php"],
    ["CUHK general Non-JUPAS", "CUHK Undergraduate Admissions - General Requirements", "https://admission.cuhk.edu.hk/application/non-jupas/general-requirements/"],
    ["HKUST Bioengineering", "HKUST Undergraduate Admissions - BEng in Bioengineering", "https://pvs0147.ust.hk/our-programs/school-of-engineering/bioengineering"],
    ["HKUST Engineering JUPAS", "HKUST School of Engineering - 2026 Admissions JUPAS", "https://seng.hkust.edu.hk/academics/undergraduate/2026-admissions/jupas"],
    ["HKUST Biomedical and Health Sciences", "HKUST BMH Programme - Admissions Routes", "https://bmh.hkust.edu.hk/admission/admission-routes"],
    ["PolyU BME", "PolyU Undergraduate Admissions - JS3150 Biomedical Engineering", "https://www.polyu.edu.hk/study/ug/jupas/2026/js3150"],
    ["PolyU general JUPAS", "PolyU - JUPAS General Entrance Requirements", "https://www.polyu.edu.hk/study/ug/admissions/jupas/jupas-general-entrance-requirements"],
    ["CityUHK BME", "CityUHK Admissions - BEng Biomedical Engineering", "https://www.cityu.edu.hk/admo/programmes/beng-biomedical-engineering"],
    ["CityUHK JUPAS listing", "JUPAS - CityUHK JS1211 BEng Biomedical Engineering", "https://www.jupas.edu.hk/en/programme/cityuhk/JS1211/"],
]

body = []
body.append(p("Undergraduate Biomedical Engineering Related Courses in Hong Kong", style="Title"))
body.append(p("Admissions requirements guide for the top five Hong Kong universities", style="Subtitle"))
body.append(p("Prepared 12 August 2026. Focus: 2026/27 undergraduate entry requirements and practical applicant planning.", style="Meta"))
body.append(p("Scope note: The top-five university set is based on the QS World University Rankings 2026 order for Hong Kong universities. The guide focuses on biomedical engineering or closely related undergraduate routes. HKUST is included through Bioengineering and Biomedical/Health Sciences routes because it does not present a programme titled exactly Biomedical Engineering in the reviewed undergraduate admissions pages.", style="BodyText"))

body.append(p("Top Five Universities", style="Heading1", keep_next=True))
body.append(table(["Rank", "Short name", "University", "Source rank"], top_rows, [700, 950, 5300, 2410]))

body.append(p("Quick Comparison", style="Heading1", keep_next=True))
body.append(table(
    ["University", "Most relevant undergraduate route", "Best fit for"],
    [[r[0], r[1], r[2]] for r in program_rows],
    [1050, 4050, 4260],
))

body.append(p("Detailed Requirements", style="Heading1"))
for row in program_rows:
    body.append(p(row[0], style="Heading2", keep_next=True))
    body.append(table(
        ["Item", "Requirement summary"],
        [
            ["Programme", row[1]],
            ["Positioning", row[2]],
            ["HKDSE/JUPAS", row[3]],
            ["Other routes and notes", row[4]],
        ],
        [1900, 7460],
    ))

body.append(p("Applicant Checklist", style="Heading1"))
for item in checklist:
    body.append(p(item, num=2))

body.append(p("How to Read This Guide", style="Heading1"))
body.append(p("Minimum levels are eligibility thresholds, not predicted offers. Where a university publishes score formulae or previous admitted-score data, use those as planning signals only. Applicants with IB, A-level, SAT/AP, national curricula, associate degrees or higher diplomas should read the relevant university route page because requirements are often qualification-specific.", style="BodyText"))
body.append(p("For Hong Kong Diploma applicants, the safest preparation pattern across these routes is English, Chinese, Mathematics, CSD, plus at least two strong science or mathematics-related electives. For engineering-oriented routes, Physics and Mathematics are particularly valuable; for biomedical and health sciences routes, Biology and Chemistry appear more often as preferred or required subjects.", style="BodyText"))

body.append(p("Sources", style="Heading1", keep_next=True))
body.append(table(["Topic", "Official or reference page", "URL"], sources, [1700, 3300, 4360]))

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 wp14"><w:body>'
    + "".join(body)
    + section_break()
    + "</w:body></w:document>"
)

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:qFormat/><w:pPr><w:spacing w:after="140"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="0B2545"/><w:sz w:val="44"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:color w:val="1F4D78"/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Meta"><w:name w:val="Meta"/><w:pPr><w:spacing w:after="180"/></w:pPr><w:rPr><w:color w:val="555555"/><w:sz w:val="19"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="200"/></w:pPr><w:rPr><w:b/><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/></w:pPr><w:rPr><w:b/><w:color w:val="2E74B5"/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="CompactTable"><w:name w:val="Compact Table"/><w:tblPr><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''

numbering_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="540"/></w:tabs><w:ind w:left="540" w:hanging="270"/><w:spacing w:after="80" w:line="300" w:lineRule="auto"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

settings_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>'''

with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/_rels/document.xml.rels", doc_rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/numbering.xml", numbering_xml)
    z.writestr("word/settings.xml", settings_xml)

print(OUT.resolve())
