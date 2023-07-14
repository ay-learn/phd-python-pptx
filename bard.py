import collections.abc  # noqa: F401

import pptx
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches
from pptx.util import Pt

# from pptx.enum.text import PP_BULLET_TYPE

# from pptx import Presentation


def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def makeParaBulletPointed(para, symbol="●"):
    """Bullets are set to Arial,
    actual text can be a different font"""
    pPr = para._p.get_or_add_pPr()
    ## Set marL and indent attributes
    pPr.set("marL", "171450")
    # pPr.set("indent", "171450")
    pPr.set("indent", "0")
    ## Add buFont
    _ = SubElement(
        parent=pPr,
        tagname="a:buFont",
        typeface="Arial",
        panose="020B0604020202020204",
        pitchFamily="34",
        charset="0",
    )
    ## Add buChar
    _ = SubElement(parent=pPr, tagname="a:buChar", char=symbol)


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
) -> None:
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


def add_paragraph(text_box=None, text=None, symbol="●") -> None:
    if text_box is None or text is None:
        exit("missing args in add_paragraph")
    paragraph = text_box.add_paragraph()
    paragraph.text = f" {text}"
    makeParaBulletPointed(paragraph, symbol)


def add_slide(master_layout, title=None, subtitle=None, symbol="●", texts=None) -> None:
    presentation = pptx.Presentation()

    presentation.slide_width = Inches(16)
    presentation.slide_height = Inches(9)

    slide = presentation.slides.add_slide(presentation.slide_layouts[master_layout])
    # TODO: if placeholder = ... do ... elif do ...elif do...

    FIT_NONE = MSO_AUTO_SIZE.NONE

    ALIGN_TOP = MSO_ANCHOR.TOP
    ALIGN_CENTER = PP_ALIGN.CENTER
    ALIGN_LEFT = PP_ALIGN.LEFT

    BLUE_DARK1 = RGBColor(30, 4, 91)
    BLUE_DARK2 = RGBColor(0, 32, 96)
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 135, 0)

    for shape in slide.placeholders:
        print("%d %s" % (shape.placeholder_format.idx, shape.name))

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
    if texts is not None:
        text_box = slide.placeholders[4]
        text_frame = text_box.text_frame
        for text in texts:
            add_paragraph(text_frame, symbol=symbol, text=text)

        set_options(
            text_box,
            text=None,
            upper=None,
            width=Inches(15),
            height=None,
            X=Inches(2.25),
            Y=Inches(1),
            color=BLUE_DARK2,
            bold=False,
            size=30,
            font="Arial",
            align_H=ALIGN_LEFT,
            align_V=ALIGN_TOP,
            shrink=FIT_NONE,
        )

    # if image is not None:
    #     image_shape = slide.shapes.add_picture(image, 0, 0, 100, 100)

    return presentation


# 1,3,7 and ... has bullet
if __name__ == "__main__":
    presentation = add_slide(
        master_layout=4,
        title="Introduction",
        subtitle="Définitions",
        symbol="-",
        texts=[
            "Une source radioactive est une quantité connue d'un radionucléide",
            "qui émet un rayonnement ionisant.",
        ],
        # image="img.png",
    )
    if presentation is not None:
        presentation.save("new_slide.pptx")
    else:
        exit("Presentation not saved")


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
