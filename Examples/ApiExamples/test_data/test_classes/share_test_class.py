# -*- coding: utf-8 -*-



class ShareTestClass:
    def __init__(self, sector, industry, ticker, weight, delta):
        self.sector = sector
        self.industry = industry
        self.ticker = ticker
        self.weight = weight
        self.delta = delta

    def title(self):
        percent_value = self.delta * 100
        return f"{Ticker}\r\n{percent_value}%"


    def color(self):
        from aspose.pydrawing import Color
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        
        full_color_delta = 0.016
        unused_color_channel_value = 80
        r = unused_color_channel_value
        g = unused_color_channel_value
        b = unused_color_channel_value
        
        # Calculate value based on Delta (assumed to be a predefined variable)
        # C#: int value = unusedColorChannelValue + (int)System.Math.Round(System.Math.Abs(Delta) / fullColorDelta * (byte.MaxValue - unusedColorChannelValue));
        value = unused_color_channel_value + int(round(abs(Delta) / full_color_delta * (255 - unused_color_channel_value)))
        
        if value > 255:
            value = 255
        
        if delta < 0:
            r = value
        else:
            g = value
        
        # Return hex color string
        return f"#{r:02X}{g:02X}{b:02X}"


    def industry_color(self):
        if self.industry == "Consumer Electronics":
            return "#1B9629"
        elif self.industry == "Software - Infrastructure":
            return "#6029E3"
        elif self.industry == "Semiconductors":
            return "#E38529"
        elif self.industry == "Internet Content & Information":
            return "#964D05"
        elif self.industry == "Entertainment":
            return "#12E32B"
        elif self.industry == "Internet Retail":
            return "#96002C"
        elif self.industry == "Auto Manufactures":
            return "#1EE3A4"
        elif self.industry == "Credit Services":
            return "#D40B70"
        else:
            return "#888888"
