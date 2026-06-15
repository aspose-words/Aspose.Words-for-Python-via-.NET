# -*- coding: utf-8 -*-

import datetime


class NumericTestBuilder:
    def __init__(self):
        self.m_value1 = 1
        self.m_value2 = 1
        self.m_value3 = 1
        self.m_value4 = 1
        self.m_logical = False
        self.m_date = datetime.datetime(2018,1,1)

    def with_values_and_date(self, value1, value2, value3, value4, date_time):
        self.m_value1 = value1
        self.m_value2 = value2
        self.m_value3 = value3
        self.m_value4 = value4
        self.m_date = date_time
        return self

    def with_values_and_logical(self, value1, value2, value3, value4, logical):
        self.m_value1 = value1
        self.m_value2 = value2
        self.m_value3 = value3
        self.m_value4 = value4
        self.m_logical = logical
        return self

    def with_values(self, value1, value2):
        self.m_value1 = value1
        self.m_value2 = value2
        return self

    def build(self):
        return NumericTestClass(self.m_value1,self.m_value2,self.m_value3,self.m_value4,self.m_logical,self.m_date)
