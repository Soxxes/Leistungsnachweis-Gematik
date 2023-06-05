import calendar
from datetime import datetime, date

import pandas as pd
import numpy as np


class Report:

    def __init__(self, group: pd.DataFrame, month: int, year: int,
                 employee_name: str, project_name: str):
        self.group = group
        self.month = month
        self.year = year
        self.employee_name = employee_name
        self.project_name = project_name

        self.references = {
            #          row, column
            "weekday": [12, 4], # D12
            "date": [12, 5], # E12
            "hours": [12, 7], # G12
            "description": [12, 14] # aka comment, N12
        }

        self.header_references = {
            "Mitarbeiter": "G3",
            "Projekt": "G5",
            "Berichtsmonat": "G7",
            "Datum": "G8"
        }

    def get_report_date(self) -> str:
        return f"{calendar.month_name[self.month]} {self.year}"
    
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
            comment = row["Comments"]
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
    def fill_header(self, sheet):
        sheet[self.header_references["Mitarbeiter"]] = self.employee_name
        sheet[self.header_references["Projekt"]] = self.project_name
        sheet[self.header_references["Berichtsmonat"]] = self.get_report_date()
        sheet[self.header_references["Datum"]] = date.today().strftime("%d.%m.%Y")
    
    # sheet is openpyxl worksheet
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