import io
import aspose.words as aw
import aspose.barcode as ab
import aspose.pydrawing as pydraw


def barcode_generator():
    #ExStart
    #ExFor:IBarcodeGenerator
    #ExFor:IBarcodeGenerator.get_barcode_image(BarcodeParameters)
    #ExFor:IBarcodeGenerator.get_old_barcode_image(BarcodeParameters)
    #ExFor:FieldOptions.barcode_generator
    #ExSummary: Shows how to use a barcode generator.
    doc = aw.Document("in.docx")
    doc.field_options.barcode_generator = CustomBarcodeGenerator()
    doc.save("out.pdf")
    #ExEnd



#ExStart
#ExFor: IBarcodeGenerator
#ExFor: IBarcodeGenerator.get_barcode_image(BarcodeParameters)
#ExFor: IBarcodeGenerator.get_old_barcode_image(BarcodeParameters)
#ExFor: FieldOptions.barcode_generator
#ExSummary: Shows how to implement a IBarcodeGenerator.
class CustomBarcodeGenerator(aw.fields.IBarcodeGenerator):

    default_qr_x_dimension_in_pixels = 4.0
    default_1_dx_dimension_in_pixels = 1.0

    def twips_to_pixels(self, heightInTwips, defVal):
        """Converts a height value in twips to pixels using a default DPI of 96."""
        return self.twips_to_pixels(heightInTwips, 96, defVal)

    def twips_to_pixels(self, heightInTwips, resolution, def_val) :
        """Converts a height value in twips to pixels based on the given resolution."""
        try :
            lVal = int(heightInTwips)
            return (lVal / 1440.0) * resolution
        except Exception as e:
            return def_val

    def get_rotation_angle(self, rotationAngle, defVal):
        """Gets the rotation angle in degrees based on the given rotation angle string."""
        match rotationAngle:
            case "0": return 0
            case "1": return 270
            case "2": return 180
            case "3": return 90
            case _ : return defVal

    def get_qr_correction_level(self, errorCorrectionLevel, def_val) -> ab.generation.QRErrorLevel:
        """Converts a string representation of an error correction level to a QRErrorLevel enum value."""
        match errorCorrectionLevel:
            case "0": return ab.generation.QRErrorLevel.LEVEL_L
            case "1": return ab.generation.QRErrorLevel.LEVEL_M
            case "2": return ab.generation.QRErrorLevel.LEVEL_Q
            case "3": return ab.generation.QRErrorLevel.LEVEL_H
            case _ :  return def_val

    def get_barcode_encode_type(self, encode_type_from_word) -> ab.generation.EncodeTypes:
        """Gets the barcode encode type based on the given encode type from Word."""
        match encode_type_from_word:
            case "QR":
                return ab.generation.EncodeTypes.QR
            case "CODE128":
                return ab.generation.EncodeTypes.CODE128
            case "CODE39":
                return ab.generation.EncodeTypes.CODE39
            case "JPPOST":
                return ab.generation.EncodeTypes.RM4SCC
            case "EAN8":
                return ab.generation.EncodeTypes.EAN8
            case "JAN8":
                return ab.generation.EncodeTypes.EAN8
            case "EAN13":
                return ab.generation.EncodeTypes.EAN13
            case "JAN13":
                return ab.generation.EncodeTypes.EAN13
            case "UPCA":
                return ab.generation.EncodeTypes.UPCA
            case "UPCE":
                return ab.generation.EncodeTypes.UPCE
            case "CASE":
                return ab.generation.EncodeTypes.ITF14
            case "ITF14":
                return ab.generation.EncodeTypes.ITF14
            case "NW7":
                return ab.generation.EncodeTypes.CODABAR
            case _:
                return ab.generation.EncodeTypes.NONE

    def convert_color(self, input_color, def_val) -> pydraw.Color:
        """Converts a hexadecimal color string to a Color object."""
        if not input_color: return def_val
        try:
            color = int(input_color, base=16)
            return pydraw.Color.from_argb((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF)
        except Exception as e:
            return def_val

    def scale_factor(self, scaleFactor, def_val):
        """Calculates the scale factor based on the provided string representation."""
        try :
            scale = int(scaleFactor)
            return scale / 100.0
        except Exception as e:
            return def_val

    def set_pos_code_style(self, gen : ab.generation.BarcodeGenerator, pos_code_style,  barcode_value):
        """Sets the position code style for a barcode generator."""
        match pos_code_style:
            # STD default and without changes.
            case "SUP2":
                gen.code_text = barcode_value[:-2]
                gen.parameters.barcode.supplement.supplement_data = barcode_value[-2:]
            case "SUP5":
                gen.code_text = barcode_value[:-5]
                gen.parameters.barcode.supplement.supplement_data = barcode_value[-5:]
            case "CASE":
                gen.parameters.border.visible = True
                gen.parameters.border.color = gen.parameters.barcode.bar_color
                gen.parameters.border.DashStyle = ab.BorderDashStyle.SOLID
                gen.parameters.border.Width.Pixels = gen.parameters.barcode.x_dimension.pixels * 5

    def draw_error_image(self, error) -> io.BytesIO:
        """Draws a simple error image with the given error text."""
        # Create a simple image with the text "error"
        from PIL import Image, ImageDraw, ImageFont

        img = Image.new("RGB", (100, 100), color=(255, 255, 255))
        d = ImageDraw.Draw(img)
        font = ImageFont.load_default()
        d.text((10, 40), error, font=font, fill=(255, 0, 0))

        # Save image to a bytes buffer and return raw bytes
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format="PNG")
        img_byte_arr.seek(0)
        return img_byte_arr

    def get_barcode_image(self, parameters: aw.fields.BarcodeParameters) -> io.BytesIO:

        try:
            gen = ab.generation.BarcodeGenerator(self.get_barcode_encode_type(parameters.barcode_type), parameters.barcode_value)

            # Set color.
            gen.parameters.barcode.bar_color = self.convert_color(parameters.foreground_color, gen.parameters.barcode.bar_color)
            gen.parameters.back_color = self.convert_color(parameters.background_color, gen.parameters.back_color)

            # Set display or hide text.
            if not parameters.display_text:
                gen.parameters.barcode.code_text_parameters.location = ab.generation.CodeLocation.NONE
            else:
                gen.parameters.barcode.code_text_parameters.location = ab.generation.CodeLocation.BELOW

            # Set QR Code error correction level.s
            gen.parameters.barcode.qr.error_level = ab.generation.QRErrorLevel.LEVEL_H
            if parameters.error_correction_level:
                gen.parameters.barcode.qr.error_level = self.get_qr_correction_level(parameters.error_correction_level, gen.parameters.barcode.qr.qr_error_level)

            # Set rotation angle.
            if parameters.symbol_rotation:
                gen.parameters.rotation_angle = self.get_rotation_angle(parameters.symbol_rotation, gen.parameters.rotation_angle)

            # Set scaling factor.
            scalingFactor = 1
            if parameters.scaling_factor:
                scalingFactor = self.scale_factor(parameters.scaling_factor, scalingFactor)

            # Set size.
            if gen.barcode_type == ab.generation.EncodeTypes.QR:
                gen.parameters.barcode.x_dimension.pixels = max(1.0, round(self.default_qr_x_dimension_in_pixels * scalingFactor))
            else:
                gen.parameters.barcode.x_dimension.pixels = max(1.0, round(self.default_1_dx_dimension_in_pixels * scalingFactor))

            # Set height.
            if parameters.symbol_height:
                gen.parameters.barcode.bar_height.pixels = max(5.0, round(self.twips_to_pixels(parameters.symbol_height, gen.parameters.barcode.bar_height.pixels) * scalingFactor))

            # Set style of a Point-of-Sale barcode.
            if parameters.pos_code_style :
                self.set_pos_code_style(gen, parameters.pos_code_style, parameters.barcode_value)

            img_byte_arr = io.BytesIO()
            gen.save(img_byte_arr, ab.generation.BarCodeImageFormat.PNG)
            img_byte_arr.seek(0)

            return img_byte_arr

        except Exception as e:
            print("Error generating barcode: " + str(e))
            return self.draw_error_image(str(e))


    def get_old_barcode_image(self, parameters: aw.fields.BarcodeParameters) -> io.BytesIO:
        raise NotImplementedError()
#ExEnd