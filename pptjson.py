import collections.abc

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
IMG_BUTTOM = ""


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

    # def colorP(self, rgb):
    #     for paragraph in self.shape.text_frame.paragraphs:
    #             paragraph.font.color.rgb = rgb
    #         # if bold is not None:
    #         #     paragraph.font.bold = bold
    #         # if font_size is not None:
    #         #     paragraph.font.size = Pt(font_size)
    #         # if font_name is not None:
    #         #     paragraph.font.name = font_name
    #     return self

    def upper(self):
        self.shape.text = self.shape.text.upper()
        return self

    def bold(self):
        print("ici bold")
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

    def add_paragraph(self, text, symbol="●"):
        paragraph = self.shape.text_frame.add_paragraph()
        paragraph.text = f"{symbol} {text}"
        return TextElement(
            paragraph
        )  # Return a TextElement object to control the individual paragraph


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

    def add_paragraph(self, text=None, symbol="●"):
        # Return the result of calling the add_paragraph method of the text_box_shape attribute
        # This will be a TextElement object that can control the individual paragraph with dot notation
        return self.paragraph_shape.add_paragraph(text, symbol)

    def set_paragraph(self, width=None, height=None, X=None, Y=None):
        if width is not None:
            self.paragraph_shape.width(width)
        if height is not None:
            self.paragraph_shape.height(height)
        if X is not None:
            self.paragraph_shape.X(Y)
        if Y is not None:
            self.paragraph_shape.Y(Y)

        return self.paragraph_shape

    def add_image(self, image_path, image_position):
        # Use the shapes attribute to add an image to the slide
        # This is similar to your original code
        if image_position is IMG_RIGHT:
            left = Inches(12)
            top = Inches(3)
            width = Inches(3.90)
            height = Inches(3)
        else:
            left = Inches(6.5)
            top = Inches(5.30)
            width = None
            height = Inches(3)

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


##2# Example usage:
##2
##2presentation = Presentation("t8.pptx")
##2
##2# Use the add_title method instead of the title method
##2# slide1 = Slide(presentation).add_title("Hello World")
##2slide1 = Slide(presentation)
##2slide1.add_title("Hello World").upper().bold().width(12).color(RED)
##2slide1.add_subtitle("This is a subtitle").bold().height(22)
##2
##2slide2 = Slide(presentation)
##2slide2.add_title("Title2").bold().width(12)
##2slide2.add_paragraph("This is the first paragraph").height(33)
##2slide2.add_paragraph("This is the second paragraph").height(33)
##2
##2slide1.save("new_slide.pptx")

import json


def add_slide_from_data(slide_data):
    slide = Slide(presentation)

    title = slide_data.get("title", "")
    subtitle = slide_data.get("subtitle", "")
    paragraphs = slide_data.get("paragraphs", [])
    image = slide_data.get("image", "")

    if title:
        (
            slide.add_title(title)
            .X(1)
            .Y(0)
            .width(16)
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
            .width(8)
            .bold()
            .color(GREEN)
            .font_size(36)
            .font_name("Monotype Corsiva")
            .align_H(ALIGN_LEFT)
            .align_V(ALIGN_TOP)
            .shrink(FIT_NONE)
        )

    if image == IMG_BUTTOM:
        img_width = 0
    else:
        img_width = 4
    text_with = 15 - img_width

    slide.set_paragraph(8, 3, 2, 5)

    for paragraph in paragraphs:
        (
            slide.add_paragraph(paragraph).width(12)
            # .height(33)
            # .X(2)
            # .Y(1)
            # .colorP(BLUE_DARK1)
            # .font_size(36)
            # .font_name("Arial")
            # .align_H(ALIGN_LEFT)
            # .align_V(ALIGN_TOP)
            # .shrink(FIT_NONE)
        )

    return slide


def load_slides_from_json(json_file_path):
    with open(json_file_path, "r") as f:
        slides_data = json.load(f)

    for slide_data in slides_data:
        add_slide_from_data(slide_data)


presentation = Presentation("t8.pptx")
load_slides_from_json("slides.json")

presentation.save("new_slide.pptx")
