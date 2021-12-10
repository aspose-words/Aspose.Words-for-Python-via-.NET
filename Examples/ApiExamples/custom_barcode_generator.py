import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class CustomBarcodeGenerator(ApiExampleBase):
    """Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)"""

    @staticmethod
    def convert_symbol_height(height_in_twips_string: str) -> float:
        """Converts barcode image height from Word units to Aspose.BarCode units."""

        # Input value is in 1/1440 inches (twips)
        height_in_twips = int(height_in_twips_string)

        # Convert to mm
        return height_in_twips * 25.4 / 1440

    @staticmethod
    def convert_color(input_color: str) -> drawing.Color:
        """Converts barcode image color from Word to Aspose.BarCode."""

        # Input should be from "0x000000" to "0xFFFFFF"
        color = int(input_color, base=16)

        return drawing.Color.from_argb(color >> 16, (color & 0xFF00) >> 8, color & 0xFF)

    @staticmethod
    def convert_scaling_factor(scaling_factor: str) -> float:
        """Converts bar code scaling factor from percent to float."""

        percent = int(scaling_factor)

        if percent < 10 or percent > 10000:
            raise Exception("Error! Incorrect scaling factor - " + scaling_factor + ".")

        return percent / 100

    def get_barcode_image(self, parameters: aw.fields.BarcodeParameters) -> drawing.Image:
        """Implementation of the get_barcode_image() method for IBarCodeGenerator interface."""

        if parameters.barcode_type is None or parameters.barcode_value is None:
            return None

        generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.QR)

        type = parameters.barcode_type.upper()

        if type == "QR":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.QR)

        elif type == "CODE128":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.CODE128)

        elif type == "CODE39":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.CODE39_STANDARD)

        elif type == "EAN8":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.EAN8)

        elif type == "EAN13":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.EAN13)

        elif type == "UPCA":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.UPCA)

        elif type == "UPCE":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.UPCE)

        elif type == "ITF14":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.ITF14)

        elif type == "CASE":
            generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.NONE)

        if generator.barcode_type == aspose.barcode.generation.EncodeTypes.NONE:
            return None

        generator.code_text = parameters.barcode_value

        if generator.barcode_type == aspose.barcode.generation.EncodeTypes.QR:
            generator.parameters.barcode.code_text_parameters.two_d_display_text = parameters.barcode_value

        if parameters.foreground_color is not None:
            generator.parameters.barcode.bar_color = convert_color(parameters.foreground_color)

        if parameters.background_color is not None:
            generator.parameters.back_color = convert_color(parameters.background_color)

        if parameters.symbol_height is not None:
            generator.parameters.image_height.pixels = convert_symbol_height(parameters.symbol_height)
            generator.parameters.auto_size_mode = aspose.barcode.generation.AutoSizeMode.NONE

        generator.parameters.barcode.code_text_parameters.location = aspose.barcode.generation.CodeLocation.NONE

        if parameters.display_text:
            generator.parameters.barcode.code_text_parameters.location = aspose.barcode.generation.CodeLocation.BELOW

        generator.parameters.caption_above.text = ""

        scale = 2.4 # Empiric scaling factor for converting Word barcode to Aspose.BarCode
        xdim = 1.0

        if generator.barcode_type == aspose.barcode.generation.EncodeTypes.QR:
            generator.parameters.auto_size_mode = aspose.barcode.generation.AutoSizeMode.NEAREST
            generator.parameters.image_width.inches *= scale
            generator.parameters.image_height.inches = generator.parameters.image_width.inches
            xdim = generator.parameters.image_height.inches / 25
            generator.parameters.barcode.x_dimension.inches = generator.parameters.barcode.bar_height.inches = xdim

        if parameters.scaling_factor is not None:
            scaling_factor = CustomBarcodeGenerator.convert_scaling_factor(parameters.scaling_factor)
            generator.parameters.image_height.inches *= scaling_factor

            if generator.barcode_type == aspose.barcode.generation.EncodeTypes.QR:
                generator.parameters.image_width.inches = generator.parameters.image_height.inches
                generator.parameters.barcode.x_dimension.inches = generator.parameters.barcode.bar_height.inches = xdim * scaling_factor

            generator.parameters.auto_size_mode = aspose.barcode.generation.AutoSizeMode.NONE

        return generator.generate_bar_code_image()

    def get_old_barcode_image(self, parameters: aw.BarcodeParameters) -> drawing.Image:
        """Implementation of the get_old_barcode_image() method for IBarCodeGenerator interface."""

        if parameters.postal_address is None:
            return None

        generator = aspose.barcode.generation.BarcodeGenerator(aspose.barcode.generation.EncodeTypes.POSTNET)
        generator.code_text = parameters.postal_address

        # Hardcode type for old-fashioned Barcode
        return generator.generate_bar_code_image()
