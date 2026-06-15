# -*- coding: utf-8 -*-



class ShareQuoteTestClass:
    def __init__(self, date, volume, open, high, low, close):
        self.date = date
        self.volume = volume
        self.open = open
        self.high = high
        self.low = low
        self.close = close

    def color(self):
        return "#1B9629" if (self.open < self.close) else "#96002C"
