# -*- coding: utf-8 -*-

import datetime


class Common:
    @staticmethod
    def get_managers():
        result = []
        manager = ManagerTestClass()
        manager.name = "John Smith"
        manager.age = 36
        init_value = ContractTestClass()
        init_value.client = ClientTestClass()
        init_value.client.name = "A Company"
        init_value.client.country = "Australia"
        init_value.client.local_address = "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"
        init_value.manager = manager
        init_value.price = 1200000
        init_value.date = datetime.datetime(2017,1,1)
        init_value2 = ContractTestClass()
        init_value2.client = ClientTestClass()
        init_value2.client.name = "B Ltd."
        init_value2.client.country = "Brazil"
        init_value2.client.local_address = "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"
        init_value2.manager = manager
        init_value2.price = 750000
        init_value2.date = datetime.datetime(2017,4,1)
        init_value3 = ContractTestClass()
        init_value3.client = ClientTestClass()
        init_value3.client.name = "C & D"
        init_value3.client.country = "Canada"
        init_value3.client.local_address = "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"
        init_value3.manager = manager
        init_value3.price = 350000
        init_value3.date = datetime.datetime(2017,7,1)
        manager.contracts = [init_value, init_value2, init_value3]
        result.append(manager)
        manager = ManagerTestClass()
        manager.name = "Tony Anderson"
        manager.age = 37
        init_value4 = ContractTestClass()
        init_value4.client = ClientTestClass()
        init_value4.client.name = "E Corp."
        init_value4.client.local_address = "445 Mount Eden Road Mount Eden Auckland 1024"
        init_value4.manager = manager
        init_value4.price = 650000
        init_value4.date = datetime.datetime(2017,2,1)
        init_value5 = ContractTestClass()
        init_value5.client = ClientTestClass()
        init_value5.client.name = "F & Partners"
        init_value5.client.local_address = "20 Greens Road Tuahiwi Kaiapoi 7691 "
        init_value5.manager = manager
        init_value5.price = 550000
        init_value5.date = datetime.datetime(2017,8,1)
        manager.contracts = [init_value4, init_value5]
        result.append(manager)
        manager = ManagerTestClass()
        manager.name = "July James"
        manager.age = 38
        init_value6 = ContractTestClass()
        init_value6.client = ClientTestClass()
        init_value6.client.name = "G & Co."
        init_value6.client.country = "Greece"
        init_value6.client.local_address = "Karkisias 6 GR-111 42  ATHINA GRÉCE"
        init_value6.manager = manager
        init_value6.price = 350000
        init_value6.date = datetime.datetime(2017,2,1)
        init_value7 = ContractTestClass()
        init_value7.client = ClientTestClass()
        init_value7.client.name = "H Group"
        init_value7.client.country = "Hungary"
        init_value7.client.local_address = "Budapest Fiktív utca 82., IV. em./28.2806"
        init_value7.manager = manager
        init_value7.price = 250000
        init_value7.date = datetime.datetime(2017,5,1)
        init_value8 = ContractTestClass()
        init_value8.client = ClientTestClass()
        init_value8.client.name = "I & Sons"
        init_value8.client.local_address = "43 Vogel Street Roslyn Palmerston North 4414"
        init_value8.manager = manager
        init_value8.price = 100000
        init_value8.date = datetime.datetime(2017,7,1)
        init_value9 = ContractTestClass()
        init_value9.client = ClientTestClass()
        init_value9.client.name = "J Ent."
        init_value9.client.country = "Japan"
        init_value9.client.local_address = "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"
        init_value9.manager = manager
        init_value9.price = 100000
        init_value9.date = datetime.datetime(2017,8,1)
        manager.contracts = [init_value6, init_value7, init_value8, init_value9]
        result.append(manager)
        return result

    @staticmethod
    def get_empty_managers():
                    return Enumerable.Empty<ManagerTestClass>();


    @staticmethod
    def get_clients():
        result = []
        for manager in Common.get_managers():
            for contract in manager.contracts:
                result.append(contract.client)
        return result

    @staticmethod
    def get_contracts():
        result = []
        for manager in Common.get_managers():
            for contract in manager.contracts:
                result.append(contract)
        return result

    @staticmethod
    def get_shares():
        return [ShareTestClass("Technology","Consumer Electronics","AAPL",6.602835,-0.0054), ShareTestClass("Technology","Software - Infrastructure","MSFT",5.832072,-0.005), ShareTestClass("Technology","Software - Infrastructure","ADBE",0.562561,-0.0274), ShareTestClass("Technology","Semiconductors","NVDA",1.335994,-0.0074), ShareTestClass("Technology","Semiconductors","QCOM",0.462198,0.0248), ShareTestClass("Communication Services","Internet Content & Information","GOOG",3.771651,0.011), ShareTestClass("Communication Services","Entertainment","DIS",0.575768,0.0102), ShareTestClass("Communication Services","Entertainment","WBD",0.116579,-0.0165), ShareTestClass("Consumer Cyclical","Internet Retail","AMZN",3.011482,0.044), ShareTestClass("Consumer Cyclical","Auto Manufactures","TSLA",1.816734,-0.0018), ShareTestClass("Consumer Cyclical","Auto Manufactures","GM",0.160205,0.0026), ShareTestClass("Financial","Credit Services","V",1.1,0.005)]

    @staticmethod
    def get_share_quotes():
        return [ShareQuoteTestClass(45131,15232450,171.32,172.5,170.69,171.98), ShareQuoteTestClass(45132,13962990,172.2,172.7,171.4,171.86), ShareQuoteTestClass(45133,14902060,171.86,171.93,170.31,171.35), ShareQuoteTestClass(45134,16962540,171.64,173.1,171.35,172), ShareQuoteTestClass(45135,15588280,171.98,172.4,170,171.44)]
