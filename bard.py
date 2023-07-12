import collections.abc

import pptx
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from pptx.util import Pt

# from pptx import Presentation


def set_options(
    shape,
    text=None,
    upper=None,
    width=None,
    X=None,
    Y=None,
    color=None,
    bold=None,
    size=None,
    font=None,
    align_H=None,
    align_V=None,
    shrink=None,
):
    if text is not None:
        shape.text = text
        if upper is not None:
            shape.text = text.upper()
    if width is not None:
        shape.width = width
    if X is not None:
        shape.top = X
    if Y is not None:
        shape.left = Y
    if color is not None:
        shape.text_frame.paragraphs[0].font.color.rgb = color
    if bold is not None:
        shape.text_frame.paragraphs[0].font.bold = bold
    if size is not None:
        shape.text_frame.paragraphs[0].font.size = Pt(size)
    if font is not None:
        shape.text_frame.paragraphs[0].font.name = font
    if align_H is not None:
        shape.text_frame.paragraphs[0].alignment = align_H
    if align_V is not None:
        shape.text_frame.vertical_anchor = align_V

    if shrink is not None:
        shape.text_frame.auto_size = shrink


def new_slide(title=None, subtitle=None, text=None, image=None):
    presentation = pptx.Presentation()

    presentation.slide_width = Inches(16)
    presentation.slide_height = Inches(9)

    slide = presentation.slides.add_slide(presentation.slide_layouts[0])

    FIT_TEXT = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    FIT_SHAPE = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    FIT_NONE = MSO_AUTO_SIZE.NONE

    ALIGN_TOP = MSO_ANCHOR.TOP
    ALIGN_CENTER = PP_ALIGN.CENTER

    BLUE_DARK1 = RGBColor(30, 4, 91)
    BLUE_DARK2 = RGBColor(0, 32, 96)
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 135, 0)

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
        subtitle_shape = slide.placeholders[1]
        set_options(
            subtitle_shape,
            text=subtitle,
            upper=None,
            width=Inches(8),
            X=Inches(1),
            Y=Inches(1),
            color=GREEN,
            bold=True,
            size=36,
            font="Monotype Corsiva",
            align_H=ALIGN_CENTER,
            align_V=ALIGN_TOP,
            shrink=FIT_NONE,
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
        subtitle="Définitions",
        text="This is the text for the slide.",
        image="img.png",
    )
    presentation.save("new_slide.pptx")
