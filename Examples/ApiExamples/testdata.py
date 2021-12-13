import io
from typing import Optional, Iterable, List
from datetime import datetime

import aspose.words as aw
import aspose.pydrawing as drawing


class ClientTestClass:
    
    def __init__(self, name: str, country: Optional[str] = None, local_address: Optional[str] = None):
        self.name = name
        self.country = country
        self.local_address = local_address


class ColorItemTestClass:

    def __init__(self, name: str, color: drawing.Color, color_code: Optional[int] = None, value1: Optional[float] = None, value2: Optional[float] = None, value3: Optional[float] = None):
        self.name = name
        self.color = color
        self.color_code = color_code
        self.value1 = value1
        self.value2 = value2
        self.value3 = value3


class ContractTestClass:

    def __init__(self, manager: 'ManagerTestClass', client: ClientTestClass, price: float, date: datetime):
        self.manager = manager
        self.client = client
        self.price = price
        self.date = date


class DocumentTestClass:

    def __init__(self, doc: Optional[aw.Document] = None,
                 doc_stream: Optional[io.BytesIO] = None,
                 doc_bytes: Optional[bytes] = b'',
                 doc_string: Optional[str] = None):
        self.document = doc
        self.document_stream = doc_stream
        self.document_bytes = doc_bytes
        self.document_string = doc_string


class ImageTestClass:

    def __init__(self, image: Optional[drawing.Image] = None,
                 image_stream: Optional[io.BytesIO] = None,
                 image_bytes: Optional[bytes] = None,
                 image_string: Optional[str] = None):
        self.image = image
        self.image_stream = image_stream
        self.image_bytes = image_bytes
        self.image_string = image_string


class ManagerTestClass:

    def __init__(self, name: str, age: int, contracts: Iterable[ContractTestClass]):
        self.name = name
        self.age = age
        self.contracts = contracts


class MessageTestClass:

    def __init__(self, name: str, message: str):
        self.name = name
        self.message = message


class NumericTestClass:

    def __init__(self, value1: Optional[int], value2: float, value3: int, value4: Optional[int], logical: Optional[bool] = None, date: Optional[datetime] = None):
        self.value1 = value1
        self.value2 = value2
        self.value3 = value3
        self.value4 = value4
        self.logical = logical
        self.date = date

    def sum(self, value1: int, value2: int) -> int:
        result = value1 + value2;
        return result


class Common:
    
    @staticmethod
    def get_managers() -> List[ManagerTestClass]:
    
        managers: List[ManagerTestClass] = []

        manager = ManagerTestClass(name="John Smith", age=36, contracts=[])
        manager.contracts = [
            ContractTestClass(
                client=ClientTestClass(
                    name="A Company",
                    country="Australia",
                    local_address="219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"),
                manager=manager,
                price=1200000.0,
                date=datetime(2017, 1, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="B Ltd.",
                    country="Brazil",
                    local_address="Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"),
                manager=manager,
                price=750000.0,
                date=datetime(2017, 4, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="C & D",
                    country="Canada",
                    local_address="101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"),
                manager=manager,
                price=350000.0,
                date=datetime(2017, 7, 1)),
            ]
        managers.append(manager)

        manager = ManagerTestClass(name="Tony Anderson", age=37, contracts=[])
        manager.contracts = [
            ContractTestClass(
                client=ClientTestClass(
                    name="E Corp.",
                    local_address="445 Mount Eden Road Mount Eden Auckland 1024"),
                manager=manager,
                price=650000.0,
                date=datetime(2017, 2, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="F & Partners",
                    local_address="20 Greens Road Tuahiwi Kaiapoi 7691 "),
                manager=manager,
                price=550000.0,
                date=datetime(2017, 8, 1)),
            ]
        managers.append(manager)

        manager = ManagerTestClass(name="July James", age=38, contracts=[])
        manager.contracts = [
            ContractTestClass(
                client=ClientTestClass(
                    name="G & Co.",
                    country="Greece",
                    local_address="Karkisias 6 GR-111 42  ATHINA GRÉCE"),
                manager=manager,
                price=350000.0,
                date=datetime(2017, 2, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="H Group",
                    country="Hungary",
                    local_address="Budapest Fiktív utca 82., IV. em./28.2806"),
                manager=manager,
                price=250000.0,
                date=datetime(2017, 5, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="I & Sons",
                    local_address="43 Vogel Street Roslyn Palmerston North 4414"),
                manager=manager,
                price=100000.0,
                date=datetime(2017, 7, 1)),
            ContractTestClass(
                client=ClientTestClass(
                    name="J Ent.",
                    country="Japan",
                    local_address="Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"),
                manager=manager,
                price=100000.0,
                date=datetime(2017, 8, 1)),
            ]
        managers.append(manager)

        return managers

    @staticmethod
    def get_empty_managers() -> List[ManagerTestClass]:
        return []

    @staticmethod
    def get_clients() -> List[ClientTestClass]:
        clients: List[ClientTestClass] = []
        for manager in Common.get_managers():
            for contract in manager.contracts:
                clients.append(contract.client)
        return clients

    @staticmethod
    def get_contracts() -> List[ContractTestClass]:
        contracts: List[ContractTestClass] = []
        for manager in Common.get_managers():
            for contract in manager.contracts:
                contracts.append(contract)
        return contracts
