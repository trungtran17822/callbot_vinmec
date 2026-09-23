from __future__ import annotations

import html
import zipfile
from pathlib import Path


OUT = Path("Callbot_Benh_vien_A_14_slides.pptx")
SLIDE_W = 12192000
SLIDE_H = 6858000

COLORS = {
    "navy": "12355B",
    "blue": "1D8CF8",
    "teal": "00B8A9",
    "green": "3BB273",
    "coral": "FF6B6B",
    "yellow": "FFD166",
    "purple": "7B61FF",
    "light": "F7FBFF",
    "mint": "E8FFF8",
    "sky": "E8F4FF",
    "ink": "203040",
    "muted": "637083",
    "white": "FFFFFF",
}


slides = [
    {
        "kicker": "Product proposal",
        "title": "Callbot cho Bệnh viện A",
        "subtitle": "Tự động hỗ trợ đặt lịch, hủy lịch và nhắc lịch khám cho bệnh nhân",
        "bullets": [
            "Giảm tải cho lễ tân tổng đài",
            "Tăng tốc độ phản hồi cho bệnh nhân",
            "Tích hợp thông báo với App A",
        ],
        "accent": "blue",
        "type": "cover",
    },
    {
        "kicker": "Bối cảnh",
        "title": "Vấn đề đang đến từ cả hai phía",
        "bullets": [
            "Lễ tân phải xử lý lượng cuộc gọi lớn trong cùng thời điểm.",
            "Bệnh nhân cần thông tin chính xác về khoa, bác sĩ, cơ sở và lịch còn trống.",
            "Thông tin xác nhận dễ bị bỏ sót khi quy trình phụ thuộc nhiều vào thao tác thủ công.",
        ],
        "accent": "teal",
    },
    {
        "kicker": "Pain points",
        "title": "Lễ tân tổng đài bị quá tải",
        "bullets": [
            "Không xử lý kịp toàn bộ cuộc gọi đến.",
            "Dễ bỏ qua hoặc ghi thiếu thông tin đã xác nhận với khách hàng.",
            "Khó nắm đầy đủ lịch khám của từng khoa, bác sĩ và cơ sở.",
            "Phải dành thời gian cho các cuộc gọi lặp lại, ít phức tạp.",
        ],
        "accent": "coral",
    },
    {
        "kicker": "Pain points",
        "title": "Bệnh nhân chờ lâu và thiếu chắc chắn",
        "bullets": [
            "Có thể phải chờ rất lâu khi tổng đài quá tải.",
            "Thông tin đăng ký có nguy cơ bị bỏ qua hoặc xác nhận thiếu.",
            "Không rõ nên khám khoa nào, ở đâu, bác sĩ nào phụ trách.",
            "Trải nghiệm đặt lịch ảnh hưởng trực tiếp đến niềm tin với bệnh viện.",
        ],
        "accent": "purple",
    },
    {
        "kicker": "Tác động",
        "title": "Vấn đề quan trọng vì xảy ra thường xuyên",
        "bullets": [
            "Chờ lâu tạo cảm giác ức chế và làm giảm sự hài lòng.",
            "Bỏ sót thông tin làm bệnh nhân mất niềm tin khi cần chăm sóc y tế.",
            "Thiếu hướng dẫn trước khi khám khiến hành trình bệnh nhân kém rõ ràng.",
            "Rủi ro mất nhóm khách hàng thân thiết và ảnh hưởng doanh thu bệnh viện.",
        ],
        "accent": "yellow",
    },
    {
        "kicker": "Hành vi hiện tại",
        "title": "Người dùng đang tự xoay sở",
        "bullets": [
            "Đến bệnh viện đặt lịch trực tiếp để tránh bị bỏ sót thông tin.",
            "Chuyển sang đăng ký khám ở bệnh viện khác.",
            "Một số trường hợp hiếm gọi trực tiếp cho bác sĩ để hỏi lịch.",
        ],
        "accent": "green",
    },
    {
        "kicker": "Khoảng trống",
        "title": "Các cách hiện tại đều có điểm yếu",
        "bullets": [
            "Đặt lịch trực tiếp vẫn phải chờ và tốn thời gian di chuyển.",
            "Chọn bệnh viện khác có thể làm bệnh nhân mất cơ hội sử dụng dịch vụ tốt hơn.",
            "Không phải bác sĩ nào cũng sẵn sàng chia sẻ số điện thoại cá nhân.",
            "Quy trình thiếu một kênh phản hồi nhất quán, nhanh và có kiểm soát.",
        ],
        "accent": "blue",
    },
    {
        "kicker": "Product vision",
        "title": "Từ Callbot đến hệ sinh thái MedicBot",
        "bullets": [
            "Giai đoạn đầu: đặt lịch, hủy lịch, nhắc lịch và gửi lưu ý trước khám.",
            "Tương lai: sàng lọc bệnh nhân và tư vấn ban đầu theo phạm vi được duyệt.",
            "Mở rộng: chăm sóc và theo dõi sau khám.",
            "Mục tiêu dài hạn: kênh hỗ trợ thông tin y tế 24/7 cho bệnh viện.",
        ],
        "accent": "teal",
    },
    {
        "kicker": "Value proposition",
        "title": "Giá trị mang lại cho bệnh viện và bệnh nhân",
        "bullets": [
            "Hỗ trợ tất cả bệnh nhân trong thời gian ngắn hơn.",
            "Giảm tình trạng bệnh nhân không biết khoa, phòng, bác sĩ hoặc cơ sở khám.",
            "Lễ tân tập trung vào các cuộc gọi cần xử lý chuyên sâu.",
            "Tăng tỷ lệ phản hồi tốt và chất lượng chăm sóc khách hàng.",
        ],
        "accent": "green",
    },
    {
        "kicker": "Kiến trúc sản phẩm",
        "title": "Hai hướng phát triển song song",
        "bullets": [
            "Rule-based: chạy đúng kịch bản, phù hợp các luồng đơn giản và dễ kiểm soát.",
            "Agent: linh hoạt hơn với hội thoại tự nhiên và truy vấn thông tin theo ngữ cảnh.",
            "Cách tiếp cận song song giúp team so sánh độ ổn định, trải nghiệm và phạm vi xử lý.",
        ],
        "accent": "purple",
    },
    {
        "kicker": "Lý do cần agent",
        "title": "Rule-based khó bao phủ tình huống thực tế",
        "bullets": [
            "Trùng họ tên và ngày sinh có thể gây nhầm định danh bệnh nhân.",
            "Bệnh nhân có thể chỉ nhớ năm sinh hoặc đọc năm sinh rút gọn.",
            "Câu trả lời theo script dễ thiếu tự nhiên.",
            "Người gọi thường hỏi thêm về khoa khám, cơ sở gần nhất hoặc lịch còn trống.",
        ],
        "accent": "coral",
    },
    {
        "kicker": "Voice AI & dữ liệu",
        "title": "Nền tảng giọng nói và quyền riêng tư",
        "bullets": [
            "Speech-to-text: team đang ưu tiên nghiên cứu model VinAI.",
            "Text-to-speech: lựa chọn ban đầu là FPT.AI do hiệu năng tốt trong khảo sát.",
            "Đã so sánh với các lựa chọn như ViettelAI, VNPTAI và Vieneu.",
            "Dữ liệu y tế là dữ liệu nhạy cảm, nên privacy là yêu cầu kỹ thuật bắt buộc.",
        ],
        "accent": "blue",
    },
    {
        "kicker": "Tích hợp",
        "title": "Callbot kết nối với App A và lịch khám",
        "bullets": [
            "Gửi xác nhận đặt lịch, hủy lịch hoặc thay đổi lịch khám qua App A.",
            "Gửi lưu ý cần thiết trước khi bệnh nhân đi khám.",
            "Truy cập lịch bác sĩ theo khoa, cơ sở và khung giờ còn trống.",
            "Với agent, có thể gửi thông báo bệnh nhân đến bác sĩ sau khi đặt lịch.",
        ],
        "accent": "teal",
    },
    {
        "kicker": "Guardrails",
        "title": "Ranh giới an toàn và bước tiếp theo",
        "bullets": [
            "Không tự chẩn đoán, khám bệnh, kê đơn hoặc kết luận y khoa.",
            "Không thay thế bác sĩ trong bất kỳ quyết định chuyên môn nào.",
            "Chỉ cung cấp nội dung y khoa đã được bệnh viện phê duyệt.",
            "Tiếp tục hoàn thiện UI, đo lường vận hành và chốt phạm vi MVP.",
        ],
        "accent": "coral",
    },
]


def esc(text: str) -> str:
    return html.escape(text, quote=True)


def alloc_id() -> int:
    value = getattr(alloc_id, "value", 2)
    alloc_id.value = value + 1
    return value


def shape_rect(x, y, cx, cy, color, alpha=None, radius=False):
    geom = "roundRect" if radius else "rect"
    alpha_xml = f'<a:alpha val="{alpha}"/>' if alpha else ""
    sid = alloc_id()
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{sid}" name="Shape {sid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
        <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="{color}">{alpha_xml}</a:srgbClr></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>"""


def text_box(x, y, cx, cy, text, size=2800, color=None, bold=False, align="l"):
    color = color or COLORS["ink"]
    b = '<a:b/>' if bold else ''
    sid = alloc_id()
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{sid}" name="Text {sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/><a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="{align}"/>
          <a:r><a:rPr lang="vi-VN" sz="{size}">{b}<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Aptos Display"/></a:rPr><a:t>{esc(text)}</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>"""


def bullet_box(x, y, cx, cy, bullets, size=2050):
    sid = alloc_id()
    paras = []
    for bullet in bullets:
        paras.append(
            f"""
        <a:p>
          <a:pPr marL="342900" indent="-228600">
            <a:buChar char="•"/>
          </a:pPr>
          <a:r><a:rPr lang="vi-VN" sz="{size}"><a:solidFill><a:srgbClr val="{COLORS['ink']}"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>{esc(bullet)}</a:t></a:r>
        </a:p>"""
        )
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{sid}" name="Bullets {sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/><a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" spcFirstLastPara="1"/>
        <a:lstStyle/>
        {''.join(paras)}
      </p:txBody>
    </p:sp>"""


def slide_xml(slide, idx):
    alloc_id.value = 2
    accent = COLORS[slide["accent"]]
    shapes = [
        shape_rect(0, 0, SLIDE_W, SLIDE_H, COLORS["light"]),
        shape_rect(0, 0, 520000, SLIDE_H, accent),
        shape_rect(9280000, 0, 2912000, 6858000, accent, alpha="15000"),
        shape_rect(9700000, 900000, 1200000, 1200000, COLORS["yellow"], alpha="45000", radius=True),
        shape_rect(10650000, 4550000, 1050000, 1050000, COLORS["teal"], alpha="35000", radius=True),
        shape_rect(915000, 760000, 1320000, 305000, accent, radius=True),
        text_box(1030000, 812000, 1100000, 220000, slide["kicker"].upper(), 1150, COLORS["white"], True, "ctr"),
    ]

    if slide.get("type") == "cover":
        shapes.extend(
            [
                text_box(920000, 1400000, 8400000, 1120000, slide["title"], 4600, COLORS["navy"], True),
                text_box(940000, 2650000, 7600000, 650000, slide["subtitle"], 2200, COLORS["muted"]),
                bullet_box(1000000, 3600000, 7100000, 1600000, slide["bullets"], 2050),
                shape_rect(8050000, 1750000, 2550000, 2550000, COLORS["white"], alpha="75000", radius=True),
                text_box(8350000, 2450000, 1950000, 700000, "MedicBot", 3000, accent, True, "ctr"),
                text_box(980000, 6120000, 5000000, 260000, "MVP: đặt lịch • hủy lịch • nhắc lịch khám", 1250, COLORS["muted"]),
            ]
        )
    else:
        shapes.extend(
            [
                text_box(900000, 1220000, 8450000, 920000, slide["title"], 3500, COLORS["navy"], True),
                shape_rect(915000, 2290000, 1700000, 65000, accent),
                bullet_box(950000, 2650000, 8500000, 2950000, slide["bullets"], 2020),
            ]
        )

    shapes.extend(
        [
            text_box(10150000, 6050000, 900000, 280000, f"{idx:02d}/14", 1200, COLORS["muted"], True, "ctr"),
            text_box(905000, 6120000, 4850000, 260000, "Callbot cho Bệnh viện A", 1150, COLORS["muted"]),
        ]
    )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {''.join(shapes)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def content_types():
    overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, len(slides) + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {overrides}
</Types>"""


def presentation_xml():
    sld_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, len(slides) + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{len(slides) + 2}"/></p:sldMasterIdLst>
  <p:sldIdLst>{sld_ids}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr><a:defRPr lang="vi-VN"/></a:defPPr>
  </p:defaultTextStyle>
</p:presentation>"""


def presentation_rels():
    rels = "\n".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, len(slides) + 1)
    )
    theme_id = len(slides) + 1
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {rels}
  <Relationship Id="rId{theme_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId{theme_id + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>"""


def root_rels():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def simple_xml_parts():
    return {
        "ppt/presProps.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>""",
        "ppt/viewProps.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr></p:viewPr>""",
        "ppt/tableStyles.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>""",
        "docProps/app.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>{len(slides)}</Slides></Properties>""",
        "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Callbot cho Bệnh viện A</dc:title><dc:creator>Codex</dc:creator><cp:keywords>callbot, hospital, MedicBot</cp:keywords><dc:description>14-slide presentation based on the provided documentation.</dc:description></cp:coreProperties>""",
        "ppt/theme/theme1.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Bright Professional"><a:themeElements><a:clrScheme name="Bright Professional"><a:dk1><a:srgbClr val="{COLORS['navy']}"/></a:dk1><a:lt1><a:srgbClr val="{COLORS['white']}"/></a:lt1><a:dk2><a:srgbClr val="{COLORS['ink']}"/></a:dk2><a:lt2><a:srgbClr val="{COLORS['light']}"/></a:lt2><a:accent1><a:srgbClr val="{COLORS['blue']}"/></a:accent1><a:accent2><a:srgbClr val="{COLORS['teal']}"/></a:accent2><a:accent3><a:srgbClr val="{COLORS['coral']}"/></a:accent3><a:accent4><a:srgbClr val="{COLORS['yellow']}"/></a:accent4><a:accent5><a:srgbClr val="{COLORS['green']}"/></a:accent5><a:accent6><a:srgbClr val="{COLORS['purple']}"/></a:accent6><a:hlink><a:srgbClr val="{COLORS['blue']}"/></a:hlink><a:folHlink><a:srgbClr val="{COLORS['purple']}"/></a:folHlink></a:clrScheme><a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>""",
        "ppt/slideMasters/slideMaster1.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>""",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>""",
        "ppt/slideLayouts/slideLayout1.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>""",
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>""",
    }


def build():
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types())
        z.writestr("_rels/.rels", root_rels())
        z.writestr("ppt/presentation.xml", presentation_xml())
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels())
        for name, data in simple_xml_parts().items():
            z.writestr(name, data)
        for i, slide in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml(slide, i))
            z.writestr(
                f"ppt/slides/_rels/slide{i}.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
            )


if __name__ == "__main__":
    build()
    print(f"Wrote {OUT.resolve()}")
