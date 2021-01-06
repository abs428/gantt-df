"""
Script that creates simple Gantt charts from pandas DataFrames.

Typical usage is as follows:
>>> import pandas as pd
>>> from gantt import gantt_to_excel
>>> # The data is read from a file in practice
>>> data = pd.DataFrame(
    {'START DATE': ['2021-01-06',
                    '2021-01-06',
                    '2021-01-07',
                    '2021-01-13'],
 'END DATE': ['2021-01-08',
            '2021-01-12',
            '2021-01-12',
            '2021-01-21'],
 'TASK': ['Define components',
        'Research',
        'Requirement gathering',
        'Build demo'],
 'DURATION (days)': [3, 5, 4, 7]}
)
>>> print(data.columns)
Index(['START DATE', 'END DATE', 'TASK', 'DURATION (days)'], dtype='object')
>>> gantt_to_excel(data = data, start_col = "START DATE", end_col = "END DATE", 
               duration_col = "DURATION (days)", description = 'TASK', output = 'out_test.xlsx')
"""

import pandas as pd
import numpy as np
import arrow
import typing
import xlsxwriter

DAYS = tuple(
    ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
)


def is_workday(
    date: str, weekend: typing.Iterable[str], holidays: typing.Iterable[str] = {}
) -> bool:
    """Helper function that returns True if the input is a workday. Optional argument `holidays` is an iterable of holidays"""
    mapping = dict(enumerate(DAYS))
    holidays = {arrow.get(day) for day in holidays}
    date = arrow.get(date)
    return (date not in holidays) and (mapping[date.weekday()] not in weekend)


def generate_date_series(
    start_date: str,
    end_date: str,
    weekend: typing.Iterable[str] = {"saturday", "sunday"},
    holidays: typing.Iterable[str] = {},
):
    """Function that generates all the dates between the start date and end dates (both inclusive)
    after ignorning the holidays

    start_date: str
        Start date as a string. Recommended format YYYY-MM-DD.

    end_date: str
        End date specified as a string. Recommended format YYYY-MM-DD.

    weekend: Iterable[str], default={'saturday', 'sunday'}
        Iterable containing the days of the week that are not workdays.

    holidays: Iterable[str], default={}
        Iterable containing the days of the week that are not workdays. Recommended format YYYY-MM-DD.

    Raises
    ------
    AssertionError:
        If weekend is not specified correctly

    Returns
    -------
    result: pd.DatetimeIndex
        A datetime index that contains the required dates
    """
    weekend = {elem.lower() for elem in weekend}
    assert weekend.issubset(DAYS), "Weekend is not specified correctly."

    date_range = pd.date_range(start_date, end_date)
    return pd.to_datetime(
        [day for day in date_range if is_workday(day, weekend, holidays)]
    )


def where(date: str, date_range: typing.Union[np.array, pd.Series]):
    """Helper function that returns the index of `date` in an iterable `date_range`.
    Assumes that `date` occurs only once in the iterable. Returns -1 if `date` is
    not found."""
    return np.cumsum(np.flip(date_range == date)).sum() - 1


def gantt_to_excel(
    data: pd.DataFrame,
    start_col: str,
    end_col: str,
    duration_col: str,
    description: str,
    output: str,
    date_format: str = "d-m-yyyy",
    colour: str = "f79646",
    symbol: str = "",
    only_workdays: bool = True,
    holidays: typing.Iterable[str] = {},
):
    """Function that converts a DataFrame into Gantt chart in an Excel spreadsheet.

    data: pd.DataFrame
        Input dataframe containing start and end dates, durations and descriptions of
        every task

    start_col: str
        Column that contains the start dates of every task

    end_col: str
        Column that contains the end dates of every task

    duration_col: str
        Column that contains the duration (in days) for every task

    output: str
        Name of the output file to be generated. Typically has the extension of .xls / .xlsx

    date_format: str, default="d-m-yyyy"
        The formatting of the dates in the output. The supported options can be found at
        https://xlsxwriter.readthedocs.io/working_with_dates_and_time.html#working-with-dates-and-time

    colour: str, default="f79646"
        Sets the colour of the bars in the generated Gantt chart. The default is a shade of
        orange. A colour pallette of reference is http://wordfaqs.ssbarnhill.com/Word%202007%20Color%20Swatches.pdf

    symbol: str, default=""
        Displays symbols in the bars populated in the Gantt chart. Blank by default. Can be
        used to simulate ASCII-like output

    only_workdays: bool, default=True
        Flag for whether only workdays are displayed on the chart

    holidays: Iterable[str], default={}
        Optional list of holidays
    """
    assert {start_col, end_col, duration_col, description}.issubset(
        data.columns
    ), "Some of the columns are not present in the data"
    assert data.notnull().any(None), "Nulls are not permitted in the data."

    data = data.copy()  # Don't mutate the original dataframe
    data[start_col] = pd.to_datetime(data[start_col])
    data[end_col] = pd.to_datetime(data[end_col])

    row_nums = {
        desc: row
        for row, desc in enumerate(
            data.groupby(description)
            .apply(lambda x: x[start_col].min())
            .sort_values()
            .index
        )
    }
    data.index = data[description].map(row_nums)

    # Setting up the workbook object
    workbook = xlsxwriter.Workbook(output)
    # Formats
    # https://xlsxwriter.readthedocs.io/working_with_dates_and_time.html#working-with-dates-and-time
    date_format = workbook.add_format({"num_format": date_format, "bold": True})
    bold_format = workbook.add_format({"bold": True})
    cell_colour = workbook.add_format()
    # Pick colours from http://wordfaqs.ssbarnhill.com/Word%202007%20Color%20Swatches.pdf
    cell_colour.set_bg_color(colour)
    worksheet = workbook.add_worksheet("Chart")

    min_date, max_date = data[start_col].min(), data[end_col].max()
    if only_workdays:
        date_range = generate_date_series(min_date, max_date, holidays=holidays)
    else:
        date_range = generate_date_series(min_date, max_date, {}, {})

    for col, day in enumerate(date_range):
        worksheet.write(0, col + 1, day, date_format)

    endpoints = zip(data[start_col], data[end_col], data.index)
    for task in data[description]:
        start, end, row = next(endpoints)
        worksheet.write(row + 1, 0, task, bold_format)
        start_index = where(start, date_range) + 1
        end_index = where(end, date_range) + 2

        for col in range(
            start_index,
            end_index,
        ):
            worksheet.write(row + 1, col, symbol, cell_colour)

    workbook.close()
