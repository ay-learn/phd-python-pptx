import collections.abc  # noqa: F401

import pptx
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
# from pptx.enum.text import PP_BULLET_TYPE
from pptx.util import Inches
from pptx.util import Pt

# from pptx import Presentation


def set_options(
    shape,
    text=None,
    upper=None,
    width=None,
    height=None,
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
    if height is not None:
        shape.height = height
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


def new_slide(placeholder,title=None, subtitle=None, text=None, image=None):
    presentation = pptx.Presentation()

    presentation.slide_width = Inches(16)
    presentation.slide_height = Inches(9)

    slide = presentation.slides.add_slide(presentation.slide_layouts[placeholder])
    # TODO: if placeholder = ... do ... elif do ...elif do...

    FIT_NONE = MSO_AUTO_SIZE.NONE

    ALIGN_TOP = MSO_ANCHOR.TOP
    ALIGN_CENTER = PP_ALIGN.CENTER
    ALIGN_LEFT   = PP_ALIGN.LEFT

    RGBColor(30, 4, 91)
    RGBColor(0, 32, 96)
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 135, 0)

    for shape in slide.placeholders:
        print('%d %s' % (shape.placeholder_format.idx, shape.name))

    if title is not None:
        title_shape = slide.shapes.title

        set_options(
            title_shape,
            text=title,
            upper=True,
            width=Inches(16),
            height=None,
            X=Inches(0.5),
            Y=Inches(0),
            color=RED,
            bold=True,
            size=36,
            font="Arial",
            align_H=ALIGN_CENTER,
            align_V=ALIGN_TOP,
            shrink=FIT_NONE,
        )


    if subtitle is not None:
        subtitle_shape = slide.placeholders[1]
        set_options(
            subtitle_shape,
            text=subtitle,
            upper=None,
            width=Inches(8),
            height=None,
            X=Inches(2),
            Y=Inches(1),
            color=GREEN,
            bold=True,
            size=36,
            font="Monotype Corsiva",
            align_H=ALIGN_LEFT,
            align_V=ALIGN_TOP,
            shrink=FIT_NONE,
        )
    if text is not None:

        slide.placeholders[1].text= 'Find the bullet slide layout'
        text_box = slide.placeholders[1].text_frame
        paragraph1 = text_box.add_paragraph()
        paragraph1.level = 1
        paragraph1.text = "This is the first paragraph"

        # from pptx.text.text import _Run, BulletFormat
        paragraph2 = text_box.add_paragraph()
        paragraph2.level = 0
        paragraph2.text = "This is the second paragraph"
        # char = "●"


        # set_options(
        #     text_box,
        #     text=text,
        #     upper=None,
        #     width=Inches(4),
        #     height=None,
        #     X=Inches(4),
        #     Y=Inches(1),
        #     color=BLUE_DARK2,
        #     bold=None,
        #     size=40,
        #     font="Arial",
        #     align_H=ALIGN_LEFT,
        #     align_V=ALIGN_TOP,
        #     shrink=FIT_TEXT,
        # )


    # if image is not None:
    #     image_shape = slide.shapes.add_picture(image, 0, 0, 100, 100)

    return presentation

# 1,3,7 and ... has bullet
if __name__ == "__main__":
    presentation = new_slide(
        placeholder=7,
        title="Introduction",
        subtitle="Définitions",
        text="Une source radioactive",
        # text="Une source radioactive est une quantité connue d'un radionucléide qui émet un rayonnement ionisant.",
        # image="img.png",
    )
    presentation.save("new_slide.pptx")


###------------------------------------------------------------
# i: 0
#       0 Title 1
#       1 Subtitle 2
# i: 1
#       0 Title 1
#       1 Content Placeholder 2
# i: 2
#       0 Title 1
#       1 Text Placeholder 2
# i: 3
#       0 Title 1
#       1 Content Placeholder 2
#       2 Content Placeholder 3
# i: 4
#       0 Title 1
#       1 Text Placeholder 2
#       2 Content Placeholder 3
#       3 Text Placeholder 4
#       4 Content Placeholder 5
# i: 5
#       0 Title 1
# i: 6
# i: 7
#       0 Title 1
#       1 Content Placeholder 2
#       2 Text Placeholder 3
# i: 8
#       0 Title 1
#       1 Picture Placeholder 2
#       2 Text Placeholder 3
# i: 9
#       0 Title 1
#       1 Vertical Text Placeholder 2
# X: 10
#       0 Vertical Title 1
#       1 Vertical Text Placeholder 2
