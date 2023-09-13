import calendar
from datetime import datetime, date
from abc import ABC, abstractmethod

import pandas as pd

from utils.utils import add_logging

"""
IMPORTANT:
since there must not be any information about a client, I use a mapping
like this: ClientReport + <id>
the client id can be found in the config file and should never be exposed
"""

class Report(ABC):

    def __init__(self,
                 group: pd.DataFrame):
        self.group = group

    @abstractmethod
    def fill_header(self, *args, **kwargs) -> None:
        pass

    @abstractmethod
    def fill_worksheet(self, *args, **kwargs) -> None:
        pass


class ClientReport1(Report):

    def __init__(self,
                 group: pd.DataFrame,
                 month: int,
                 year: int,
                 employee_name: str,
                 project_name: str,
                 references: dict,
                 header_references: dict):
        super().__init__(group)
        
        self.month = month
        self.year = year
        self.employee_name = employee_name
        self.project_name = project_name

        # cell references in the template and output file
        self.references = references if references else None
        self.header_references = header_references if header_references else None

    def get_report_date(self) -> str:
        if self.month and self.year:
            return f"{calendar.month_name[self.month]} {self.year}"
        raise Exception("No month or year (or both not) provided.")
    
    def get_weekdays_to_dates(self) -> tuple[list[str], list[str]]:
        cal = calendar.monthcalendar(self.year, self.month)
        weekday_names = calendar.weekheader(3).split()

        weekdays = []
        dates = []
        for week in cal:
            for i, day in enumerate(week):
                if day != 0:
                    weekdays.append(weekday_names[i])
                    date = datetime(self.year, self.month, day).strftime("%d.%m.%Y")
                    dates.append(date)

        return weekdays, dates
    
    def get_hours_by_date(self) -> dict:
        d = {}
        for _, row in self.group.iterrows():
            date = row["Entry Date"].to_pydatetime().strftime("%d.%m.%Y")
            hours = row["Hours"]
            if not date in d.keys():
                d[date] = hours
            # if, for whatever reason, employee booked multiple times on the same date
            else:
                # print(f"[WARNING] {self.employee_name} booked multiple times on same date: {date}")
                d[date] += hours
        return d

    # code_to_activity should consists of a mapping from a code, e.g. "001", to
    # a specific activity, e.g. "Gematik-Abstimmung"
    # additional_comments is a list of exceptions for the codes where the actual comment
    # should be used as well instead of just the activity corresponding to the code
    def get_comment_by_date(self, code_to_activity: dict, additional_comments: list) -> dict:
        d = {}
        for _, row in self.group.iterrows():
            date = row["Entry Date"].to_pydatetime().strftime("%d.%m.%Y")
            comment = str(row["Comments"])
            if not comment:
                continue
            for code, activity in code_to_activity.items():
                if code in comment:
                    # only if the code is not in the additional comments list
                    # replace the comment by the code
                    if code not in additional_comments:
                        comment = activity
                    break
            if not date in d.keys():
                d[date] = comment
            else:
                # take longer comment in case of multiple occasions of same date
                d[date] = d[date] if len(comment) < len(d[date]) else comment
        return d
    
    # sheet is openpyxl worksheet
    @add_logging
    def fill_header(self, sheet):
        if self.header_references is not None:
            sheet[self.header_references["Mitarbeiter"]] = self.employee_name
            sheet[self.header_references["Projekt"]] = self.project_name
            sheet[self.header_references["Berichtsmonat"]] = self.get_report_date()
            sheet[self.header_references["Datum"]] = date.today().strftime("%d.%m.%Y")
    
    # sheet is openpyxl worksheet
    @add_logging
    def fill_worksheet(self, sheet, code_to_activity, additional_comments):
        weekdays, dates = self.get_weekdays_to_dates()
        hours = self.get_hours_by_date()
        comments = self.get_comment_by_date(code_to_activity, additional_comments)
        for i in range(0, len(weekdays)):
            sheet.cell(
                row=self.references["weekday"][0] + i,
                column=self.references["weekday"][1],
                value=weekdays[i]
            )
            sheet.cell(
                row=self.references["date"][0] + i,
                column=self.references["date"][1],
                value=dates[i]
            )
            if dates[i] in hours.keys():
                # same row as in dates can be used
                sheet.cell(
                    row=self.references["date"][0] + i,
                    column=self.references["hours"][1],
                    value=hours.get(dates[i])
                )
            if dates[i] in comments.keys():
                # same row as in dates can be used
                sheet.cell(
                    row=self.references["date"][0] + i,
                    column=self.references["description"][1],
                    value=comments.get(dates[i])
                )


class ClientReport2(Report):

    def __init__(self,
                 group: pd.DataFrame,
                 task_name: str,
                 grades: dict,
                 header_references: dict):
        super().__init__(group)

        self.task_name = task_name
        self.grades = grades
        self.header_references = header_references

    @add_logging
    def fill_header(self, sheet) -> None:
        ref = self.header_references[self.task_name]
        sheet[ref] = self.group["Hours"].sum()

    @add_logging
    def fill_worksheet(self, sheet, code_to_activity, additional_comments) -> None:
        # delete ugly formatting from excel sheet, start at 2 to keep header
        sheet.delete_rows(2, sheet.max_row)
        for _, row in self.group.iterrows():
            date = row["Entry Date"]
            employee_name = row["First Name"].strip() + " " + row["Last Name"].strip()
            grade = self.grades[employee_name]
            hours = row["Hours"]

            comment = self._get_comment(row["Comments"], code_to_activity, additional_comments)
            
            info = [grade, date, hours, comment]
            sheet.append(info)

    def _get_comment(self, raw_comment, code_to_activity, additional_comments) -> str:
        comment = raw_comment
        # someone wrote 1, 2, 5, etc. instead of 001, 002, 005
        if isinstance(comment, int):
            comment = "00" + str(comment)
        elif pd.isna(comment):
            comment = "MISSING COMMENT"
        for code, activity in code_to_activity.items():
            if code in comment:
                # only if the code is not in the additional comments list
                # replace the comment by the code
                if code not in additional_comments:
                    comment = activity
                break
        return comment
    