#!/usr/bin/env python3
import collections.abc
import json
import os
import random
import string
import sys
import xml.etree.ElementTree as ET

import cairosvg
import cv2
import pptx
import yaml
from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches
from pptx.util import Pt

SHRINK_NONE = MSO_AUTO_SIZE.NONE
SHRINK_TEXT = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
SHRINK_SHAPE = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

ALIGN_V_TOP = MSO_VERTICAL_ANCHOR.TOP
ALIGN_V_MIDDLE = MSO_VERTICAL_ANCHOR.MIDDLE
ALIGN_V_BOTTOM = MSO_VERTICAL_ANCHOR.BOTTOM
ALIGN_V_MIXED = MSO_VERTICAL_ANCHOR.MIXED

MSO_VERTICAL_ANCHOR.MIDDLE

ALIGN_H_CENTER = PP_ALIGN.CENTER
ALIGN_H_LEFT = PP_ALIGN.LEFT
ALIGN_H_RIGHT = PP_ALIGN.RIGHT
ALIGN_H_JUSTIFY = PP_ALIGN.JUSTIFY_LOW
ALIGN_H_DISTRIBUTE = PP_ALIGN.DISTRIBUTE
ALIGN_H_THAI_DISTRIBUTE = PP_ALIGN.THAI_DISTRIBUTE
ALIGN_H_MIXED = PP_ALIGN.MIXED

BLUE_DARK1 = RGBColor(30, 4, 91)
BLUE_DARK2 = RGBColor(0, 32, 96)
RED = RGBColor(255, 0, 0)
GREEN = RGBColor(0, 135, 0)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)

IMG_BUTTOM = "IMG_BUTTOM"
IMG_RIGHT = "IMG_RIGHT"


def get_image_dpi(image_path):
    try:
        with Image.open(image_path) as img:
            dpi = img.info.get("dpi")
            print("image_path", image_path, "DPI: ", dpi)
            if dpi:
                return dpi
            else:
                return None
    except Exception as e:
        return None


def generate_random_string(length=6):
    letters = string.ascii_lowercase
    return "".join(random.choice(letters) for _ in range(length))


# Define a base class for text elements
class TextElement:
    def __init__(self, shape):
        self.shape = shape

    def text(self, text):
        self.shape.text = text
        return self

    def width(self, inches):
        self.shape.width = inches
        return self

    def height(self, inches):
        self.shape.height = inches
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

    def spacing(self):
        # for paragraph in self.shape.text_frame.paragraphs:
        #     # help(paragraph)
        #     paragraph.line_spacing = Pt(100)
        #     paragraph.space_before = Pt(50)
        #     paragraph.space_after = Pt(200)

        return self

    def align_H(self, alignment):
        for paragraph in self.shape.text_frame.paragraphs:
            paragraph.alignment = alignment
        return self

    def align_V(self, anchor):
        self.shape.text_frame.vertical_anchor = anchor
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

    def add_paragraph_1(self, text, color=None, font_size=None, font_name=None):
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

    def add_paragraph_2(self, text, color=None, font_size=None, font_name=None):
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

    def add_paragraph_3(self, text, color=None, font_size=None, font_name=None):
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


def get_img_width(image_path, image_position):
    if image_path is None or image_path == "":
        print(f"image_path={image_path}", file=sys.stderr)
        return Inches(0), image_path
    if not os.path.isfile(image_path):
        print(f"File {image_path} does not found", file=sys.stderr)
        return Inches(0), image_path

    if image_position is None:
        image_position == IMG_BUTTOM

    if image_path.lower().endswith(".svg"):
        rand_name = generate_random_string(6)
        output_path = f"./build/{rand_name}.png"
        cairosvg.svg2png(url=image_path, write_to=output_path)
        image_path = output_path

    image_np = cv2.imread(image_path)

    image_width, image_height = 0, 0
    try:
        image_height, image_width, _ = image_np.shape
    except AttributeError as e:
        print("image_path=", image_path, "\n", e)
        return Inches(0), image_path

    # PIL way
    # image = Image.open(image_path)
    # image_width, image_height = image.size

    xy = get_image_dpi(image_path)
    if xy is not None:
        dpi_x, dpi_y = xy[0], xy[1]
    else:
        dpi_x, dpi_y = 100, 100
    image_width_inches = image_width / dpi_x
    image_height_inches = image_height / dpi_y

    width = Inches(4) * image_width_inches / image_height_inches
    if image_position == IMG_RIGHT:
        if width > Inches(5):
            width = Inches(5)
    else:
        if width > Inches(14):
            width = Inches(14)

    return width, image_path


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
        self.subtitle_shape_1 = Subtitle(self.slide.placeholders[1])
        self.subtitle_shape_2 = Subtitle(self.slide.placeholders[5])
        self.page_shape_3 = Subtitle(self.slide.placeholders[6])
        self.paragraph_shape_1 = Paragraph(self.slide.placeholders[2])
        self.paragraph_shape_2 = Paragraph(self.slide.placeholders[3])
        self.paragraph_shape_3 = Paragraph(self.slide.placeholders[4])

        # Create an attribute for the shapes collection
        self.shapes = self.slide.shapes

    # Define methods to access the attributes of the elements
    def add_title(self, text=None):  # Rename this method
        if text is not None:
            self.title_shape.text(text)

        # Return the title_shape attribute to use its methods with dot notation
        return self.title_shape

    def add_subtitle_1(self, text=None):
        if text is not None:
            self.subtitle_shape_1.text(text)
        return self.subtitle_shape_1

    def add_subtitle_2(self, text=None):
        if text is not None:
            self.subtitle_shape_2.text(text)
        return self.subtitle_shape_2

    def add_page(self, text=None):
        if text is not None:
            self.page_shape_3.text(text)
        return self.page_shape_3

    # paragraph.font.color.rgb = RED

    def add_paragraph_1(self, text=None, color=None, font_size=None, font_name=None):
        return self.paragraph_shape_1.add_paragraph_1(text, color, font_size, font_name)

    def add_paragraph_2(self, text=None, color=None, font_size=None, font_name=None):
        return self.paragraph_shape_2.add_paragraph_2(text, color, font_size, font_name)

    def add_paragraph_3(self, text=None, color=None, font_size=None, font_name=None):
        return self.paragraph_shape_3.add_paragraph_3(text, color, font_size, font_name)

    def set_paragraph_1(self, width=None, height=None, X=None, Y=None, rgb=None):
        if width is not None:
            self.paragraph_shape_1.width(width)
        if height is not None:
            self.paragraph_shape_1.height(height)
        if X is not None:
            self.paragraph_shape_1.X(Y)
        if Y is not None:
            self.paragraph_shape_1.Y(Y)
        if rgb is not None:
            self.paragraph_shape_1.color(rgb)
        return self.paragraph_shape_1

    def set_paragraph_2(self, width=None, height=None, X=None, Y=None, rgb=None):
        if width is not None:
            self.paragraph_shape_2.width(width)
        if height is not None:
            self.paragraph_shape_2.height(height)
        if X is not None:
            self.paragraph_shape_2.X(Y)
        if Y is not None:
            self.paragraph_shape_2.Y(Y)
        if rgb is not None:
            self.paragraph_shape_2.color(rgb)
        return self.paragraph_shape_2

    def set_paragraph_3(self, width=None, height=None, X=None, Y=None, rgb=None):
        if width is not None:
            self.paragraph_shape_3.width(width)
        if height is not None:
            self.paragraph_shape_3.height(height)
        if X is not None:
            self.paragraph_shape_3.X(Y)
        if Y is not None:
            self.paragraph_shape_3.Y(Y)
        if rgb is not None:
            self.paragraph_shape_3.color(rgb)
        return self.paragraph_shape_3

    def add_image(self, image_path=None, image_position=None):
        if image_path is None or image_path == "":
            print(f"image_path={image_path}", file=sys.stderr)
            return self
        if not os.path.isfile(image_path):
            print(f"File {image_path} does not found", file=sys.stderr)
            return self

        height = Inches(4)
        width, image_path = get_img_width(image_path, image_position)

        if image_position == IMG_RIGHT:
            left = Inches(16) - width
            top = Inches(3)

            self.shapes.add_picture(image_path, left=left, top=top, width=width)

        else:
            left = (Inches(16) - width) / 2
            top = Inches(4.30)

            self.shapes.add_picture(
                image_path, left=left, top=top, width=width, height=height
            )
        return self

    def save(self, path):
        if self.presentation is not None:
            self.presentation.save(path)
            print(f"---------\nsave => {path}")
        else:
            exit("Presentation not saved")


def Presentation(path):
    return pptx.Presentation(path)


def remove_first_slide(presentation):
    xml_slides = presentation.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[0])


def load_slides_from_json(json_file_path):
    with open(json_file_path, "r") as f:
        slides_data = json.load(f)

    for slide_data in slides_data:
        add_slide_from_data(slide_data)


def load_slides_from_yaml(yaml_file_path):
    with open(yaml_file_path, "r") as f:
        slides_data = yaml.safe_load(f)

    for slide_data in slides_data:
        add_slide_from_data(slide_data)


def add_slide_from_data(slide_data):
    slide = Slide(presentation)

    title = slide_data.get("title", "")
    subtitle_1 = slide_data.get("subtitle_1", "")
    subtitle_2 = slide_data.get("subtitle_2", "")
    page = slide_data.get("subtitle_3", "")
    paragraphs_1 = slide_data.get("paragraphs_1", [])
    paragraphs_2 = slide_data.get("paragraphs_2", [])
    paragraphs_3 = slide_data.get("paragraphs_3", [])
    image = slide_data.get("image_path", "")
    image_position = slide_data.get("image_position", IMG_BUTTOM)

    img_width, _ = get_img_width(image, image_position)

    if image_position == IMG_RIGHT:
        par_width = int(Inches(15) - img_width)
        if par_width < 0:
            print("img_width is bigger")
            par_width = 0
        print(
            "IMG_RIGHT: 15 Inches(",
            Inches(15),
            ") - img_width(",
            img_width,
            ")",
            end="",
        )
        print("\tpar_width:", par_width)

    else:
        par_width = Inches(15)

    # '''

    if title:
        (
            slide.add_title(title)
            .X(0.5)
            .Y(0.5)
            .width(Inches(15))
            .height(Inches(4))
            .upper()
            .color(RED)
            .bold()
            .font_size(36)
            .font_name("Arial")
            .align_H(ALIGN_H_CENTER)
            .align_V(ALIGN_V_TOP)
            .shrink(SHRINK_SHAPE)
        )

    if subtitle_1:
        (
            slide.add_subtitle_1(subtitle_1)
            .X(2)
            .Y(1)
            .width(Inches(14))
            .height(Inches(20))
            .bold()
            .color(GREEN)
            .font_size(36)
            .font_name("Monotype Corsiva")
            .align_H(ALIGN_H_LEFT)
            .align_V(ALIGN_V_TOP)
            .shrink(SHRINK_TEXT)
        )

    # if subtitle_2:
    #     (
    #         slide.add_subtitle_2(subtitle_2)
    #         .X(2)
    #         .Y(1)
    #         .width(Inches(14))
    #         .height(Inches(20))
    #         .bold()
    #         .color(GREEN)
    #         .font_size(36)
    #         .font_name("Monotype Corsiva")
    #         .align_H(ALIGN_H_LEFT)
    #         .align_V(ALIGN_V_TOP)
    #         .shrink(SHRINK_TEXT)
    #     )
    if page:
        (
            slide.add_page(page)
            .X(8.35)
            .Y(7.50)
            .width(Inches(14))
            .height(Inches(20))
            # .bold()
            .color(WHITE)
            .font_size(18)
            .font_name("Arial")
            .align_H(ALIGN_H_LEFT)
            .align_V(ALIGN_V_TOP)
            .shrink(SHRINK_TEXT)
        )

    (
        slide.set_paragraph_1()
        .width(par_width)
        .height(Inches(20))
        .X(2.5)
        .Y(1)
        .align_H(ALIGN_H_RIGHT)
        .spacing()
    )
    # (
    #     slide.set_paragraph_2()
    #     .width(par_width)
    #     .height(Inches(20))
    #     .X(2.5)
    #     .Y(1)
    #     .align_H(ALIGN_H_RIGHT)
    #     .spacing()
    # )
    (
        slide.set_paragraph_3()
        .width(par_width)
        .height(Inches(1))
        .X(4.5)
        .Y(2)
        .align_H(ALIGN_H_RIGHT)
        .spacing()
    )

    p_size = 22
    if paragraphs_1:
        for paragraph in paragraphs_1:
            slide.add_paragraph_1(
                text=paragraph,
                color=BLACK,
                font_size=p_size,
                font_name="Arial",
            )

    # if paragraphs_2:
    #     for paragraph in paragraphs_2:
    #         slide.add_paragraph_2(
    #             text=paragraph,
    #             color=BLACK,
    #             font_size=p_size,
    #             font_name="Arial",
    #         )
    #
    # if paragraphs_3:
    #     for paragraph in paragraphs_3:
    #         slide.add_paragraph_3(
    #             text=paragraph,
    #             color=BLACK,
    #             font_size=p_size,
    #             font_name="Arial",
    #         )

    slide.add_image(image_path=image, image_position=image_position)
    # '''
    return slide


presentation = Presentation("t2.pptx")

# load_slides_from_json("ppt1.json")
load_slides_from_yaml("ppt1.yaml")

remove_first_slide(presentation)
presentation.save("new_slide.pptx")
