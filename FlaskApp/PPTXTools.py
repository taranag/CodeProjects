import pptx
from pptx.util import Pt
from pptx.util import Inches

logoLeft, logoTop, logoHeight, logoWidth = 8293608, 18288, 768096, 841248
logoPath = "SeekLogo.png"

def setTableFontSize(table, fontSize):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(fontSize)

def createTitleSlide(prs, title, startDate, endDate):
    """
    Creates a blank slide with a title and logo image
    """
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    title = title.replace("\n", " ")

    titleShape = slide.shapes.title
    titleShape.text = title

    subtitleShape = slide.shapes[1]
    subtitleShape.text = "From " + startDate + " to " + endDate

    titleTextFrame = slide.shapes[0].text_frame
    titleTextFrame.paragraphs[0].runs[0].font.bold = True

    logo = slide.shapes.add_picture(logoPath, pptx.util.Inches(8.45), logoTop, pptx.util.Inches(1.54), pptx.util.Inches(1.42))
    return slide

def createBlankSlideWithTitle(prs, title, fontSize=44):
    """
    Creates a blank slide with a title and logo image
    """
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    title = title.replace("\r\n", " ")

    titleShape = slide.shapes.title
    titleShape.text = title

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame

    text_frame.paragraphs[0].runs[0].font.size = Pt(fontSize)
    text_frame.paragraphs[0].runs[0].font.bold = True

    left = top = width = height = pptx.util.Inches(1)
    slide.shapes.add_picture(logoPath, logoLeft, logoTop, logoWidth, logoHeight)
    return slide

def createTableWithDownloadHeaders(slide, rows, cols, left, top, width, height):
    """
    Creates a table on a slide.
    """
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    table.columns[2].width = Inches(2)
    table.columns[0].width = Inches(2)
    table.columns[3].width = Inches(1)
    table.rows[0].cells[1].text = "Downloaded"
    table.rows[0].cells[2].text = "Yet to Download"
    table.rows[0].cells[3].text = "Total"
    table.rows[0].cells[4].text = "% Download"
    return table

def createTableWithLearnHeaders(slide, headers, rows, cols):
    """
    Creates a table on a slide.
    """
    left = Inches(.2)
    top = Inches(1.5)
    width = Inches(9.6)
    height = Inches(2)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    # table.table.columns[2].width = Inches(2)
    # table.table.columns[0].width = Inches(2)
    #table.table.columns[cols-1].width = Inches(1)
    for i in range(len(headers)):
        table.rows[0].cells[i+1].text = headers[i]
    return table

def createTableWithFinalLearnHeaders(slide, headers, rows, cols):
    """
    Creates a table on a slide.
    """
    left = Inches(.2)
    top = Inches(1.5)
    width = Inches(10)
    height = Inches(2)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    # table.table.columns[2].width = Inches(2)
    # table.table.columns[0].width = Inches(2)
    #table.table.columns[cols-1].width = Inches(1)
    for i in range(len(headers)):
        table.rows[0].cells[i+1].text = headers[i]
    table.rows[0].cells[len(headers)+1].text = "Not on Learn"
    table.rows[0].cells[len(headers)+2].text = "Total"
    table.columns[len(headers)+2].width = Inches(1)
    return table