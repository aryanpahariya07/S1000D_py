from lxml import etree
def find_image_target(image_id, rels_file):
    with open(rels_file, "rb") as f:
        tree = etree.parse(f)
        root = tree.getroot()
        namespaces = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
        for rel in root.findall(".//rel:Relationship", namespaces=namespaces):
            if rel.get("Id") == image_id:
                return rel.get("Target")
    return None

def process_table(table_element, tr_elements, namespaces):
    table = etree.Element("table")
    tgroup = etree.Element("tgroup")
    tbody = etree.Element("tbody")
    thead = etree.Element("thead")
    max_columns = 0
    first_row = True

    for tr in tr_elements:
        table_row = etree.Element("row")

        for tc in tr.findall(".//w:tc", namespaces=namespaces):
            entry = etree.Element("entry")
            para = etree.Element("para")
            text = ""
            for t in tc.findall(".//w:t", namespaces=namespaces):
                if t.text is not None:
                    text += t.text

            para.text = text.strip() if text else ""
            entry.append(para)
            table_row.append(entry)

        if table_row.getchildren():
            if first_row:
                thead.append(table_row)
                first_row = False
            else:
                tbody.append(table_row)

        table_row_columns = len(table_row.getchildren())
        if table_row_columns > max_columns:
            max_columns = table_row_columns

    tgroup.append(thead)
    tgroup.append(tbody)
    tgroup.set("cols", str(max_columns))
    table_element.append(tgroup)

docx_xml_file = "uploads/word/document.xml"
s1000d_xml_file = "s1000d.xml"

s1000d_doc = etree.Element("s1000d")
namespaces = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

tree = etree.parse(docx_xml_file)
root = tree.getroot()
body_element = root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")

levelled_para1 = None
levelled_para2 = None
list_type = None

for element in body_element:
    if element.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p":
        p = element
        p_style = p.find("w:pPr/w:pStyle", namespaces=namespaces)
        drawing = p.find(".//w:drawing", namespaces=namespaces)

        if p_style is not None and p_style.get("{%s}val" % namespaces["w"]) == "Heading1":
            if levelled_para2 is not None:
                levelled_para1.append(levelled_para2)
            if levelled_para1 is not None:
                s1000d_doc.append(levelled_para1)

            levelled_para1 = etree.Element("LevelledPara")
            text = ""
            for t in p.iter("{%s}t" % namespaces["w"]):
                if t.text is not None:
                    text += t.text

            title = etree.Element("title")
            title.text = text
            title.set("style", "1")
            levelled_para1.append(title)

        elif p_style is not None and p_style.get("{%s}val" % namespaces["w"]) == "Heading2":
            if levelled_para2 is not None:
                levelled_para1.append(levelled_para2)

            levelled_para2 = etree.Element("LevelledPara")

            text = ""
            for t in p.iter("{%s}t" % namespaces["w"]):
                if t.text is not None:
                    text += t.text

            title = etree.Element("title")
            title.text = text
            title.set("style", "2")
            levelled_para2.append(title)

        elif p_style is not None and p_style.get("{%s}val" % namespaces["w"]) == "Title":
            text = ""
            for t in p.iter("{%s}t" % namespaces["w"]):
                if t.text is not None:
                    text += t.text

            title = etree.Element("title")
            title.text = text
            s1000d_doc.append(title)

        elif drawing is not None:
            figure = etree.Element("figure")
            graphic = etree.Element("graphic")
            blip = drawing.find(".//a:blip", namespaces=namespaces)

            if blip is not None and "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed" in blip.attrib:
                image_id = blip.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"]
                image_rels_file = "uploads/word/_rels/document.xml.rels"
                image_target = find_image_target(image_id, image_rels_file)

                if image_target is not None:
                    image_file_path = f"uploads/{image_target}"
                    graphic.set("infoEntityIdent", image_file_path)

            figure.append(graphic)
            if levelled_para2 is not None:
                levelled_para2.append(figure)
            elif levelled_para1 is not None:
                levelled_para1.append(figure)

        else:
            num_id = p.find("w:pPr/w:numPr/w:numId", namespaces=namespaces)
            if num_id is not None:
                val = num_id.get("{%s}val" % namespaces["w"])
                if val == "1" or val == "2":
                    if list_type is None or list_type != val:
                        if val == "1":
                            list_type = "1"
                            list_element = etree.Element("sequencelist")
                        else:
                            list_type = "2"
                            list_element = etree.Element("randomlist")

                        levelled_para1.append(list_element)

                    text = ""
                    for t in p.iter("{%s}t" % namespaces["w"]):
                        if t.text is not None:
                            text += t.text

                    if text:
                        listitem = etree.Element("listitem")
                        para = etree.Element("para")
                        para.text = text
                        listitem.append(para)
                        list_element.append(listitem)

            else:
                if levelled_para2 is not None:
                    text = ""
                    for t in p.iter("{%s}t" % namespaces["w"]):
                        if t.text is not None:
                            text += t.text

                    if text:
                        para = etree.Element("para")
                        para.text = text
                        levelled_para2.append(para)

                elif levelled_para1 is not None:
                    text = ""
                    for t in p.iter("{%s}t" % namespaces["w"]):
                        if t.text is not None:
                            text += t.text
                    if text:
                        para = etree.Element("para")
                        para.text = text
                        levelled_para1.append(para)

    elif element.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl":
        table_element = etree.Element("table")
        tr_elements = element.findall(".//w:tr", namespaces=namespaces)
        process_table(table_element, tr_elements, namespaces)
        if levelled_para2 is not None:
            levelled_para2.append(table_element)
        elif levelled_para1 is not None:
            levelled_para1.append(table_element)

if levelled_para2 is not None:
    levelled_para1.append(levelled_para2)
if levelled_para1 is not None:
    s1000d_doc.append(levelled_para1)

s1000d_tree = etree.ElementTree(s1000d_doc)
s1000d_tree.write(s1000d_xml_file, encoding="utf-8", xml_declaration=True, pretty_print=True)