import datetime as dt
import os
import numpy as np
import pandas as pd
import xlwings as xw
from xlwings import pro
from sqlalchemy import *
from sqlalchemy.engine import create_engine
from sqlalchemy import text
import logging
import json

try:
    engine = create_engine('bigquery://', credentials_info=os.environ["GCP_SA"])
except Exception as e:
    logging.error("NO DB CONNECTION!")
    logging.error(f"Error: {e}")


@pro.func
def print_client_engine():
    try:
        engine = create_engine('bigquery://', credentials_info=json.load(os.environ["GCP_SA"]))
        message = f"Engine: {engine}"
    except Exception as e:
        message = f"Error: {e}"
    return message

# SAMPLE 1: Hello World
@pro.func
def hello(name):
    return f"Hello {name}!"

# SAMPLE 2: Numpy, Namespace, doc strings, dynamic arrays
# This sample also shows how you set a namespace and add description to the
# function and its arguments. The namespace makes this function turn up as
# NUMPY.STANDARD_NORMAL in Excel. Multi-dimensional arrays are automatically
# spilled via Excel's native dynamic arrays, no code change required.

@pro.func(namespace="numpy")
@pro.arg("rows", doc="the number of rows in the returned array.")
@pro.arg("columns", doc="the number of columns in the returned array.")
def standard_normal(rows, columns):
    """Returns an array of standard normally distributed random numbers"""
    rng = np.random.default_rng()
    return rng.standard_normal(size=(rows, columns))


# SAMPLE 3: Force a dimensionality of the input arguments via ndim
# If your argument size can go from single cells to 1- and 2-dimensional ranges,
# force it to be always 2-dimensional. Note that you don't have to do that for
# pandas DataFrame as they are always 2-dimensional by definition.
# This sample wouldn't work for single cells and 1-dimensional ranges if
# ndim=2 is left away

@pro.func
@pro.arg("values", ndim=2)
def add_one(values):
    return [[cell + 1 for cell in row] for row in values]


# SAMPLE 4: pandas DataFrame as argument and return value


@pro.func(namespace="pandas")
@pro.arg("df", pd.DataFrame, index=False, header=False)
@pro.ret(index=False, header=False)
def correl(df):
    """Like CORREL, but it works on whole matrices instead of just 2 arrays.
    Set index and header to True if your dataset has labels
    """
    return df.corr()


# SAMPLE 5: DateTime
# This sample shows how you can convert date-formatted cells to datetime objects in
# Python by either using a decorator or by using xw.to_datetime().
# On the other hand, when returning datetime objects, xlwings takes care of formatting
# the cell automatically via data types.


@pro.func(namespace="pandas")
@pro.arg("start", dt.datetime, doc="A date-formatted cell")
@pro.arg("end", doc="A date-formatted cell")
def random_timeseries(start, end):
    # Instead of using the dt.datetime converter in the decorator, you can also convert
    # a date-formatted cell to a datetime object by using xw.to_datetime(). This is
    # especially useful if you have more than one cell that needs to be transformed.
    # xlwings returns datetime objects automatically formatted in Excel via data types
    # if your version of Excel supports them.
    end = xw.to_datetime(end)
    date_range = pd.date_range(start, end)
    rng = np.random.default_rng()
    data = rng.standard_normal(size=(len(date_range), 1))
    return pd.DataFrame(data, columns=["Growth"], index=date_range)


# SAMPLE 6: DateTime within pandas DataFrames
# pandas DataFrames allow you to use parse_dates in the same way as it works with
# pd.read_csv().


@pro.func(namespace="pandas")
@pro.arg("df", pd.DataFrame, parse_dates=[0])
def timeseries_start(df):
    """Returns the earliest date of a timeseries. Expects the leftmost column to contain
    date-formatted cells in Excel (you could use the output of random_timeseries as
    input for this function).
    """
    return df.index.min()


# SAMPLE 7: Volatile functions
# Volatile functions are calculated everytime Excel calculates something, even if none
# of the cells arguments change.


@pro.func(volatile=True)
def last_calculated():
    return f"Last calculated: {dt.datetime.now()}"


@pro.func
def layoffs_fyi():
    sql = '''  
    SELECT 
    date(date) as date,
    company,
    employees_laid_off,
    concat(cast(round(percent_laid_off*100,2) as string),"%") as percent_laid_off,
    datamachine_load_time
    FROM `datamachine-407200.macro.layoffs_fyi`
    WHERE datamachine_load_time = (select max(datamachine_load_time) from `datamachine-407200.macro.layoffs_fyi`)
    ORDER BY date desc
    '''
    with engine.begin() as conn:
        df = conn.execute(text(sql)).fetchall()
    return pd.DataFrame(df)