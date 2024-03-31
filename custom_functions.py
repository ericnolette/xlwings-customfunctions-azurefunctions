import datetime as dt
import os
import numpy as np
import pandas as pd
import xlwings as xw
from xlwings import pro
from google.oauth2 import service_account
from google.cloud import bigquery


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



# Pull data from GCP
credentials = service_account.Credentials.from_service_account_info(
    text_ = {   "type": "service_account",   "project_id": "datamachine-407200",   "private_key_id": "414d2e4e95eeee8ca69747ce46f2b9c3c1a7015c",   "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCeBRYhWxEfSCo8\n3+KKsiLdzSmOhvMZlRqtYrTzV2MWwQupx03z2VY8552KkVUjj2OUhPqgjed9mky2\nblwSwrg4s0/bBvHtZHj4eJhiXevVsjLwqt5ZbX1ZMf/yIcay3yXdgwxbjl/dOmgR\nN9NdPX5oG0+IuevoIKxpRuAXWUJGBrRN0/xWF0vvb9Vyp5DuCb+YmLbB+v2Ei9Rl\ndOHEtEvfctEvPduaDvQLE5ARfuTUhFYBCWYlneRQLttD+ooAgrMPpDN59yuRue8o\nYYgEm8Lj//M6i5PoFHU+aabcpiHDVnaU3RB8B8ZNt4WPQZP8KxNr8HIt4T6rDylg\nUjl1wW3PAgMBAAECggEARf2yw6t2dgiczMHNsekdEGkjEwIrAxXL7yhdRbtbJGT1\nKYohuPR3AhsE6K9eqpWEYfBUonC4LCX//H39QkIFHvgtUrJMzf1Gp6eit08qekin\nz3mfarTYZH2FBFZ5kBjOyMKW4qa5R1/rYFT9xbrLFomiPMk8+GAgfbDq6OycMi9r\nU3hZTxtTVac3vgVva+JyZK6uhYQJYxn6mIfGVXvUUHChglzzb47DYr/4wQM+lwD5\nTGeD/a+JJRNRvF1wIkdf/tCSIug0+1SmSfLn0pGEZLh8ovqI9u6IQUnTjOrVJy1O\nBe4uiBPeVZkRQoLsIXD0/mBpOT0B+5auPv/o2ftYwQKBgQDZHfvyJrkpiZPe/ABq\nVhrWRImWmxvdLJB8KYlzQTYDkm66wgG62U/jOqJ/fM3Y5RRErJXmCjReacWX1u1k\nag8nnJMpZ/ij7XCGSP+CEYhIIU6ZJ6ChTKT102+kbgE6BFoGPbl8eywW5hj3/E+e\neu9Lmf+R2M/jfsLa8GsP9r1LiQKBgQC6UbZGBLaM7HN2SzeqO6JOeR/+ozuzdBE8\nfBoBo9KTrxNhXMZBBrcROWTvIVaRA01aaQKT2eiHKiM3FQJQBw69hcIMHvSMMktw\n/51psqZYB2c5crM/Nt4GhxLNJAyQYl0583tOYiEEX1WuqrhCTE87SQVyJmCqAqGE\niT9jr3LglwKBgQC4/RownRvAr27lW7OC5rBhBe5w+uGH1jOZBs8M+2/pJTfhOfG9\nYPD3O3s+wnilJ7HYPXBOmz05gEeR8tc7aj5VUsv0SJkKGwF3+PRyvzttsatFRQVQ\nyXv309nYsL2s0A5gKPFEhbHwJMb7a+fusPH4aVLe0mt2ewfNAXFHHcT1eQKBgAYx\n31CWqYcn+XLOb2xejTf0uQabYMnHqycKrUaurrqwUIGlNwZEdePBt8RnpFwv8ut1\noFtQHHYaBY+4SBpnEatlfh0vDkx3A6EfLpmsEfHNVTZIxQLuDRXEefCOKUjHrHfX\ny5rAkn51uQCUtomlxeCfvemcswwUCFDCy3PCCpzDAoGBALIpq8nIKdpukk/7SFhp\n4+l1oUgW8XRGv6hpYX4Y5pBsjCJi+hYz9y8UGN2ExAYKSYfzteQRkscPlgQ/J+Go\neJkzIfNVNepY7gI5n5qc5DVXU6G/6Bm9KZEWAMoDiBj54PkXotnwtEoRcKQG5Vt9\nlDXXvnxzpH1j//x9BSZFZktt\n-----END PRIVATE KEY-----\n",   "client_email": "streamlit-datamachine@datamachine-407200.iam.gserviceaccount.com",   "client_id": "101039704290156623568",   "auth_uri": "https://accounts.google.com/o/oauth2/auth",   "token_uri": "https://oauth2.googleapis.com/token",   "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",   "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/streamlit-datamachine%40datamachine-407200.iam.gserviceaccount.com",   "universe_domain": "googleapis.com" }
)


@pro.func(namespace="pandas")
def layoffs_fyi():
    sql = """
    SELECT 
    date(date) as date,
    company,
    employees_laid_off,
    concat(cast(round(percent_laid_off*100,2) as string),"%") as percent_laid_off,
    datamachine_load_time
    FROM `datamachine-407200.macro.layoffs_fyi`
    WHERE datamachine_load_time = (select max(datamachine_load_time) from `datamachine-407200.macro.layoffs_fyi`)
    ORDER BY date desc
    """
    df = client.query(sql).to_dataframe()
    return df