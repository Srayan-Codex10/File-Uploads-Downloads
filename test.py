import requests
from io import BytesIO
from PIL import Image
from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.shared import RGBColor, Inches
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import docx
from haggis.files.docx import list_number
from docx.enum.dml import MSO_THEME_COLOR_INDEX
import traceback
import typing

# import lxml

# HTML Content
html_content = open("test.html", "r").read()  # Replace with the actual HTML content

# Parse HTML with BeautifulSoup
soup = BeautifulSoup(html_content, "lxml")

# Create a Word document
doc = Document()
doc.core_properties.title = soup.title.string if soup.title else "HTML to DOCX"
doc.add_heading(doc.core_properties.title, level=0)


# Function to convert CSS color to hex code
def extract_hex_color(style: str) -> dict[str, str]:
    color_dict = {}
    try:
        if "color:" in style:
            color = style.split("color:")[1].split(";")[0].strip()
            if (
                color.startswith("#") and len(color) == 7
            ):  # Ensure it's a valid hex code
                color_dict["color"] = color.lstrip("#")  # Remove the '#' for RGBColor
        if "background-color:" in style:
            color = style.split("background-color:")[1].split(";")[0].strip()
            if (
                color.startswith("#") and len(color) == 7
            ):  # Ensure it's a valid hex code
                color_dict["background-color"] = color.lstrip(
                    "#"
                )  # Remove the '#' for RGBColor
    except IndexError:
        pass
    return color_dict  # Return None if no valid color is found


# Function to add styled text
def add_styled_text(paragraph, text, color=None, is_bg_color=False):
    run = paragraph.add_run(text)
    if color:
        try:
            text_color = color.get("color", "000000")
            rgb = tuple(int(text_color[i : i + 2], 16) for i in (0, 2, 4))
            run.font.color.rgb = RGBColor(*rgb)
            if is_bg_color:
                bg_color = color.get("background-color", "FFFFFF")
                # rgb = tuple(int(bg_color[i:i + 2], 16) for i in (0, 2, 4))
                highlight = parse_xml(
                    r'<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), bg_color)
                )
                parent_element = run._element.getparent()
                parent_element.insert(parent_element.index(run._element) + 1, highlight)
        except ValueError:
            pass  # Skip if the color value is invalid


# Function to handle nested lists
def process_list(list_tag, parent_paragraph=None, level=0):
    for li in list_tag.find_all("li", recursive=False):
        for child in li.children:
            if child.name != "ol" and child.name != "ul":
                paragraph = doc.add_paragraph(
                    style="List Number" if list_tag.name == "ol" else "List Bullet"
                )
                paragraph.paragraph_format.left_indent = Inches(level * 0.5)
            if child.name == "a":  # Handle anchor tags
                add_styled_text(paragraph, child.text.strip())
            elif child.name == "span":
                span_color = extract_hex_color(child.get("style", ""))
                add_styled_text(paragraph, child.text.strip(), span_color)
            elif child.name == "p":
                add_styled_text(paragraph, child)
                if paragraph.style.name == "List Number":
                    list_number(doc, paragraph, prev=parent_paragraph)
                parent_paragraph = paragraph
            elif child.name == "ol" or child.name == "ul":
                process_list(child, parent_paragraph=None, level=level + 1)


# Function to process tables
""" def process_table(table_tag):
    rows = table_tag.find_all('tr')
    table = doc.add_table(rows=0, cols=len(rows[0].find_all(['th', 'td'])))
    table.style = 'Table Grid'
    
    for row in rows:
        cells = row.find_all(['th', 'td'])
        row_cells = table.add_row().cells
        for i, cell in enumerate(cells):
            cell_text = cell.text.strip()
            row_cells[i].text = cell_text
            # Apply background color to header cells
            if cell.name == 'th' and 'style' in cell.attrs:
                bg_color = extract_hex_color(cell['style'])
                if bg_color:
                    rgb = tuple(int(bg_color[i:i + 2], 16) for i in (0, 2, 4))
                    shading = row_cells[i]._tc.get_or_add_tcPr().add_new_shd()
                    shading.set_val('clear')
                    shading.set_fill(bg_color) """


# Function to add an image
def add_image(doc: Document, img_tag: Tag):
    if "src" in img_tag.attrs:
        img_url = img_tag["src"]
        try:
            response = requests.get(img_url)
            response.raise_for_status()
            image = Image.open(BytesIO(response.content))
            with BytesIO() as img_buffer:
                image.thumbnail((300, 300))  # Resize image
                image.save(img_buffer, format=image.format)
                img_buffer.seek(0)
                doc.add_picture(img_buffer, width=Inches(2))
        except Exception as e:
            traceback.print_exc()
            print(f"Failed to load image: {img_url}, Error: {e}")
            paragraph = doc.add_paragraph()
            paragraph.add_run(f"[Image not available: {img_url}]")


def add_hyperlink(paragraph, text, url):
    """
    A function that adds a hyperlink to a paragraph.
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.ns.qn('r:id'), r_id, )

    # Create a w:r element and a new w:rPr element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run()
    r._r.append(hyperlink)

    # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True


def set_paragraph_id(paragraph: docx.text.paragraph.Paragraph, p_id: str) -> None:
    """Sets paragraph ID using correct Word XML namespace"""
    try:
        # Create paragraph properties element if it doesn't exist
        if not paragraph._element.pPr:
            paragraph._element.get_or_add_pPr()
            
        # Create paragraphId tag with proper namespace
        para_id = parse_xml(f'<w:paraId w:val="{p_id}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        paragraph._element.pPr.append(para_id)
    except Exception as e:
        print(f"Failed to set paragraph ID: {str(e)}")


# Process the HTML content
bg_color = False
for tag in soup.body.descendants:
    if tag.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
        h_level = int(tag.name[1])
        doc.add_heading(tag.text.strip(), level=h_level)
    elif tag.name == "p" and tag.parent.name != "li":
        paragraph = doc.add_paragraph()
        p_id = tag.get("id", "")
        if p_id:
            set_paragraph_id(paragraph, p_id)
        color = {}
        if "style" in tag.attrs:
            color = extract_hex_color(tag["style"])
        for child in tag.children:
            if child.name == "a":  # Handle anchor tags
                # add_styled_text(paragraph, child.text.strip(), color)
                add_hyperlink(paragraph, child.text.strip(), child["href"])
            elif child.name == "span":
                span_color = extract_hex_color(child.get("style", ""))
                bg_color = True if "background-color" in child["style"] else False
                add_styled_text(
                    paragraph, child.text.strip(), span_color, is_bg_color=bg_color
                )
            elif child.name == "img":  # Handle image tags
                add_image(paragraph, child)
            elif isinstance(child, str):
                add_styled_text(paragraph, child, color)
    elif tag.name == "table":
        # process_table(tag)
        pass
    elif tag.name == "ol" or tag.name == "ul" and tag.parent.name != "li":
        process_list(tag)
    elif tag.name == "img":
        add_image(doc, tag)

# Save the document
doc.save("output.docx")
print("Document created successfully!")
