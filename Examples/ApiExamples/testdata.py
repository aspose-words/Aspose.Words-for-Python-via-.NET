# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
from typing import Optional, Iterable, List
from datetime import datetime

import aspose.words as aw
import aspose.pydrawing as drawing


class ClientTestClass:

    def __init__(self, name: str, country: Optional[str] = None, local_address: Optional[str] = None):
        self.Name = name
        self.Country = country
        self.LocalAddress = local_address


class ColorItemTestClass:

    def __init__(self, name: str, color: drawing.Color, color_code: Optional[int] = None, value1: Optional[float] = None, value2: Optional[float] = None, value3: Optional[float] = None):
        self.Name = name
        self.Color = color
        self.ColorCode = color_code
        self.Value1 = value1
        self.Value2 = value2
        self.Value3 = value3


class ContractTestClass:

    def __init__(self, client: ClientTestClass, price: float, date: datetime):
        self.Client = client
        self.Price = price
        self.Date = date


class DocumentTestClass:

    def __init__(self, doc: Optional[aw.Document] = None,
                 doc_stream: Optional[io.BytesIO] = None,
                 doc_bytes: Optional[bytes] = None,
                 doc_string: Optional[str] = None):
        self.Document = doc
        self.DocumentStream = doc_stream
        self.DocumentBytes = doc_bytes
        self.DocumentString = doc_string


class ImageTestClass:

    def __init__(self, image: Optional[drawing.Image] = None,
                 image_stream: Optional[io.BytesIO] = None,
                 image_bytes: Optional[bytes] = None,
                 image_string: Optional[str] = None):
        self.Image = image
        self.ImageStream = image_stream
        self.ImageBytes = image_bytes
        self.ImageString = image_string


class ManagerTestClass:

    def __init__(self, name: str, age: int, contracts: Iterable[ContractTestClass]):
        self.Name = name
        self.Age = age
        self.Contracts = contracts


class MessageTestClass:

    def __init__(self, name: str, message: str):
        self.Name = name
        self.Message = message


class NumericTestClass:

    def __init__(self, value1: Optional[int], value2: float, value3: int, value4: Optional[int], logical: Optional[bool] = None, date: Optional[datetime] = None):
        self.Value1 = value1
        self.Value2 = value2
        self.Value3 = value3
        self.Value4 = value4
        self.Logical = logical
        self.Date = date

    def sum(self, value1: int, value2: int) -> int:
        result = value1 + value2
        return result


class Common:

    @staticmethod
    def get_managers() -> List[ManagerTestClass]:
        return [
            ManagerTestClass(name="John Smith", age=36, contracts=[
                ContractTestClass(
                    client=ClientTestClass(
                        name="A Company",
                        country="Australia",
                        local_address="219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"),
                    price=1200000.0,
                    date=datetime(2017, 1, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="B Ltd.",
                        country="Brazil",
                        local_address="Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"),
                    price=750000.0,
                    date=datetime(2017, 4, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="C & D",
                        country="Canada",
                        local_address="101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"),
                    price=350000.0,
                    date=datetime(2017, 7, 1)),
                ]),
            ManagerTestClass(name="Tony Anderson", age=37, contracts=[
                ContractTestClass(
                    client=ClientTestClass(
                        name="E Corp.",
                        local_address="445 Mount Eden Road Mount Eden Auckland 1024"),
                    price=650000.0,
                    date=datetime(2017, 2, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="F & Partners",
                        local_address="20 Greens Road Tuahiwi Kaiapoi 7691 "),
                    price=550000.0,
                    date=datetime(2017, 8, 1)),
                ]),
            ManagerTestClass(name="July James", age=38, contracts=[
                ContractTestClass(
                    client=ClientTestClass(
                        name="G & Co.",
                        country="Greece",
                        local_address="Karkisias 6 GR-111 42  ATHINA GRÉCE"),
                    price=350000.0,
                    date=datetime(2017, 2, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="H Group",
                        country="Hungary",
                        local_address="Budapest Fiktív utca 82., IV. em./28.2806"),
                    price=250000.0,
                    date=datetime(2017, 5, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="I & Sons",
                        local_address="43 Vogel Street Roslyn Palmerston North 4414"),
                    price=100000.0,
                    date=datetime(2017, 7, 1)),
                ContractTestClass(
                    client=ClientTestClass(
                        name="J Ent.",
                        country="Japan",
                        local_address="Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"),
                    price=100000.0,
                    date=datetime(2017, 8, 1)),
                ])]

    @staticmethod
    def get_empty_managers() -> List[ManagerTestClass]:
        return []

    @staticmethod
    def get_clients() -> List[ClientTestClass]:
        clients: List[ClientTestClass] = []
        for manager in Common.get_managers():
            for contract in manager.Contracts:
                clients.append(contract.client)
        return clients

    @staticmethod
    def get_contracts() -> List[ContractTestClass]:
        contracts: List[ContractTestClass] = []
        for manager in Common.get_managers():
            for contract in manager.Contracts:
                contracts.append(contract)
        return contracts
