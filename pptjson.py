#!/usr/bin/env python3
import sys
import collections.abc
import json


import pptx
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches
from pptx.util import Pt

FIT_NONE = MSO_AUTO_SIZE.NONE

ALIGN_TOP = MSO_ANCHOR.TOP
ALIGN_CENTER = PP_ALIGN.CENTER
ALIGN_LEFT = PP_ALIGN.LEFT

BLUE_DARK1 = RGBColor(30, 4, 91)
BLUE_DARK2 = RGBColor(0, 32, 96)
RED = RGBColor(255, 0, 0)
GREEN = RGBColor(0, 135, 0)
BLACK = RGBColor(0, 0, 0)

IMG_BUTTOM = "IMG_BUTTOM"
IMG_RIGHT = "IMG_RIGHT"



# Define a base class for text elements
class TextElement:
    def __init__(self, shape):
        self.shape = shape

    def text(self, text):
        self.shape.text = text
        return self

    def width(self, inches):
        self.shape.width = Inches(inches)
        return self

    def height(self, inches):
        self.shape.height = Inches(inches)
        return self

    def X(self, inches):
        self.shape.top = Inches(inches)
        return self

    def Y(self, inches):
        self.shape.left = Inches(inches)
        return self

    def color(self, rgb):
        for paragraph in self.shape.text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.color.rgb = rgb
        return self

    def upper(self):
        self.shape.text = self.shape.text.upper()
        return self

    def bold(self):
        for paragraph in self.shape.text_frame.paragraphs:
            paragraph.font.bold = True
        return self

    def font_size(self, points):
        for paragraph in self.shape.text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.size = Pt(points)
        return self

    def font_name(self, name):
        for paragraph in self.shape.text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.name = name
        return self

    def align_H(self, alignment):
        for paragraph in self.shape.text_frame.paragraphs:
            paragraph.alignment = alignment
        return self

    def align_V(self, anchor):
        for paragraph in self.shape.text_frame.paragraphs:
            paragraph.vertical_anchor = anchor
        return self

    def shrink(self, auto_size):
        for paragraph in self.shape.text_frame.paragraphs:
            paragraph.auto_size = auto_size
        return self


# Define a subclass for title element
class Subtitle(TextElement):
    def __init__(self, shape):
        super().__init__(shape)


class Title(TextElement):
    def __init__(self, shape):
        super().__init__(shape)


class Paragraph(TextElement):
    def __init__(self, shape):
        super().__init__(shape)

    # def add_paragraph2(self, text, color=None, font_size=None,font_name=None):
    #     paragraph = self.shape.text_frame.add_paragraph()
    #     paragraph.text = text
    #
    #     paragraph.font.font_name = font_name
    #     paragraph.font.font_size = font_size
    #     paragraph.font.color.rgb = color
    #
    #     return TextElement(paragraph)

    def add_paragraph(self, text, color=None, font_size=None, font_name=None):
        paragraph = self.shape.text_frame.add_paragraph()

        run = paragraph.add_run()
        run.text = text

        if color:
            run.font.color.rgb = color
        if font_size:
            run.font.size = Pt(font_size)
        if font_name:
            run.font.name = font_name

        return TextElement(paragraph)


# p = text_frame.paragraphs[0]
# run = p.add_run()
# run.text = 'Spam, eggs, and spam'
#
# font = run.font
# font.name = 'Calibri'
# font.size = Pt(18)
# font.bold = True


class TextBox(TextElement):
    def __init__(self, shape):
        super().__init__(shape)


class Slide:
    def __init__(self, presentation):
        self.presentation = presentation

        self.presentation.slide_width = Inches(16)
        self.presentation.slide_height = Inches(9)

        self.slide = presentation.slides.add_slide(presentation.slide_layouts[0])

        # Create attributes for each element as instances of their respective classes
        self.title_shape = Title(self.slide.shapes.title)
        self.subtitle_shape = Subtitle(self.slide.placeholders[1])
        self.paragraph_shape = Paragraph(self.slide.placeholders[2])

        # Create an attribute for the shapes collection
        self.shapes = self.slide.shapes

    # Define methods to access the attributes of the elements
    def add_title(self, text=None):  # Rename this method
        if text is not None:
            self.title_shape.text(text)

        # Return the title_shape attribute to use its methods with dot notation
        return self.title_shape

    def add_subtitle(self, text=None):
        if text is not None:
            self.subtitle_shape.text(text)

        # Return the subtitle_shape attribute to use its methods with dot notation
        return self.subtitle_shape

    # paragraph.font.color.rgb = RED

    def add_paragraph(self, text=None, color=None, font_size=None, font_name=None):
        return self.paragraph_shape.add_paragraph(text, color, font_size, font_name)

    def set_paragraph(self, width=None, height=None, X=None, Y=None, rgb=None):
        if width is not None:
            self.paragraph_shape.width(width)
        if height is not None:
            self.paragraph_shape.height(height)
        if X is not None:
            self.paragraph_shape.X(Y)
        if Y is not None:
            self.paragraph_shape.Y(Y)
        if rgb is not None:
            self.paragraph_shape.color(rgb)

        return self.paragraph_shape

    def add_image(self, image_path: None, image_position):
        if image_path is not None:
            height = Inches(3)
            if image_position == IMG_RIGHT:
                left = Inches(12)
                top = Inches(3)
                width = Inches(3.90)
            else:
                left = Inches(6.5)
                top = Inches(5.30)
                width = None

            self.shapes.add_picture(image_path, left, top, width=width, height=height)
        return self

    def save(self, path):
        if self.presentation is not None:
            self.presentation.save(path)
            print(f"---------\nsave => {path}")
        else:
            exit("Presentation not saved")


def Presentation(path):
    return pptx.Presentation(path)


def load_slides_from_json(json_file_path):
    with open(json_file_path, "r") as f:
        slides_data = json.load(f)

    for slide_data in slides_data:
        add_slide_from_data(slide_data)


def add_slide_from_data(slide_data):
    slide = Slide(presentation)

    title = slide_data.get("title", "")
    subtitle = slide_data.get("subtitle", "")
    paragraphs = slide_data.get("paragraphs", [])
    image = slide_data.get("image_path", "")
    image_position = slide_data.get("image_position", IMG_BUTTOM)

    if image_position == IMG_RIGHT:
        img_width = 4
    else:
        img_width = 0

    if title:
        (
            slide.add_title(title)
            .X(1)
            .Y(0)
            .width(16)
            .height(8)
            .upper()
            .color(RED)
            .bold()
            .font_size(36)
            .font_name("Arial")
            .align_H(ALIGN_CENTER)
            .align_V(ALIGN_TOP)
            .shrink(FIT_NONE)
        )

    if subtitle:
        (
            slide.add_subtitle(subtitle)
            .X(2)
            .Y(1)
            .width(14)
            .bold()
            .color(GREEN)
            .font_size(36)
            .font_name("Monotype Corsiva")
            .align_H(ALIGN_LEFT)
            .align_V(ALIGN_TOP)
            .shrink(FIT_NONE)
        )

    (slide.set_paragraph().width(15 - img_width).X(2.5).Y(1))

    for paragraph in paragraphs:
        slide.add_paragraph(
            text=paragraph,
            color=BLACK,
            font_size=24,
            font_name="Arial",
        )

    # TODO not trow error if not has an image
    #slide.add_image(image_path=image, image_position=image_position)

    return slide


presentation = Presentation("t8.pptx")
#load_slides_from_json("slides.json")
load_slides_from_json("/tmp/ppt3.json")

presentation.save("new_slide.pptx")
