import dataclasses
from datetime import datetime


class CorrectHoursError(Exception):
    pass


@dataclasses.dataclass
class UnsupportedReportType(CorrectHoursError):
    report_type: str

    def __str__(self) -> str:
        return f"Report not supported: {self.report_type}"


@dataclasses.dataclass
class RateNotFound(CorrectHoursError):
    rate_label: str
    date: datetime

    def __str__(self) -> str:
        return (
            f"Rate not found for label \"{self.rate_label}\" and date \"{self.date.date()}\". "
            "Make sure the rates file has an entry for this combination. "
            "Also make sure the entry in the rate file doesn't have any trailing spaces."
        )
