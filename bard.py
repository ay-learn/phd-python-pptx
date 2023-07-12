import collections.abc

import pptx
# from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import PP_ALIGN
# from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches
from pptx.util import Pt

def set_options(shape,text=None, color=None, bold=None, size=None, alignment=None, anchor=None):
    # if color is not None, change the font color
    if text is not None:
        shape.text = text
    if color is not None:
        shape.text_frame.paragraphs[0].font.color.rgb = color
    # if bold is not None, change the font boldness
    if bold is not None:
        shape.text_frame.paragraphs[0].font.bold = bold
    # if size is not None, change the font size
    if size is not None:
        shape.text_frame.paragraphs[0].font.size = Pt(size)
    # if alignment is not None, change the text alignment
    if alignment is not None:
        shape.text_frame.paragraphs[0].alignment = alignment
    # if anchor is not None, change the vertical anchor
    if anchor is not None:
        shape.text_frame.vertical_anchor = anchor

def new_slide(title=None, subtitle=None, text=None, image=None):
    presentation = pptx.Presentation()

    presentation.slide_width = Inches(16)
    presentation.slide_height = Inches(9)

    slide = presentation.slides.add_slide(presentation.slide_layouts[0])

    BLUE_DARK = RGBColor(30, 4, 91)
    RED = RGBColor(255, 0, 0)

    if title is not None:
        title_shape = slide.shapes.title
        title_shape.text = title.upper()

        title_shape.width = Inches(16)
        title_shape.left = Inches(0)
        title_shape.top = Inches(1)

        title_shape.text_frame.paragraphs[0].font.color.rgb = RED

        title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP

        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.name = "Arial"

    if subtitle is not None:
        # # Monotype Corsiva
        subtitle_shape = slide.placeholders[1]
        subtitle_shape.text = subtitle
        #
        # subtitle_shape.text_frame.paragraphs[0].font.size = Pt(30)
        # subtitle_shape.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        set_options(
            subtitle_shape,
            text=subtitle,
            color=RED,
            bold=True,
            size=24,
            alignment=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.TOP,
        )
    if text is not None:
        text_box = slide.shapes.add_textbox(0, 0, 100, 100)
        text_box.text = text

    if image is not None:
        image_shape = slide.shapes.add_picture(image, 0, 0, 100, 100)

    return presentation


if __name__ == "__main__":
    presentation = new_slide(
        title="Chapter 2",
        subtitle="Introduction",
        text="This is the text for the slide.",
        image="img.png",
    )
    presentation.save("new_slide.pptx")
