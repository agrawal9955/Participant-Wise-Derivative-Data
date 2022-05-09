from dataclasses import dataclass


@dataclass
class ParticipantWiseNseRawRecords:
    dateHeader: list[str]
    columnHeader: list[str]
    clientData: list[int or str]
    diiData: list[int or str]
    fiiData: list[int or str]
    proData: list[int or str]
    totalData: list[int or str]


@dataclass
class NseHistoricalData:
    open: float
    close: float
    high: float
    low: float
