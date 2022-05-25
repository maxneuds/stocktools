from datetime import datetime as dt
import requests
import pandas as pd
import csv
import os
from time import sleep
from sys import exit

# install guide
# python -m venv env
# source env/bin/activate
# pip install -r req.txt
# manual (not using req.txt):
# pip install requests
# pip install pandas
# pip install xlsxwriter

# page: https://companiesmarketcap.com
# target base link: https://companiesmarketcap.com/?download=csv


def logger(msg):
  dt_now = dt.now().strftime("%H:%M:%S")
  print(f"{dt_now}: {msg}")


def web_get(url_target):
  logger(f"Crawling: {url_target}")
  requests.get(url_target)
  request_result = requests.get(url_target)
  return(request_result)


def csvreader_to_pandas(csv_reader):
  # read header and fix header
  header = next(csv_reader)
  header = [name.replace(" ", ".").lower() for name in header]
  # get df
  df = pd.DataFrame(csv_reader, columns=header)
  return(df)


def scraper():
  dir_output = "data"
  # create output dir
  if not os.path.exists(dir_output):
    os.makedirs(dir_output)

  # get raw csv
  url_target = "https://companiesmarketcap.com/?download=csv"
  raw_marketcap_csv = web_get(url_target)
  str_marketcap_csv = raw_marketcap_csv.content.decode('utf-8')
  csv_reader = csv.reader(str_marketcap_csv.splitlines(), delimiter=",")
  # skip first line (empty)
  next(csv_reader)

  # get dataframe
  df = csvreader_to_pandas(csv_reader)

  # export xlsx
  ts = dt.now().strftime("%y%m%d")
  path_xlsx = os.path.join(dir_output, f"{ts}-marketcap.xlsx")
  logger(f"Writing: {path_xlsx}")
  writer = pd.ExcelWriter(path_xlsx, engine='xlsxwriter')
  df.to_excel(
      excel_writer=writer,
      sheet_name="marketcap",
      freeze_panes=(1, 1),
      index=False,
      engine="xlsxwriter"
  )
  # define workbook and worksheet for formatting
  worksheet = writer.sheets["marketcap"]
  # iterate each column and set the width == the max length in that column.
  # A padding length of 2 is also added.
  for i, col in enumerate(df.columns):
    # find length of column i
    column_len = df[col].astype(str).str.len().max()
    # Setting the length if the column header is larger
    # than the max column value length
    column_len = max(column_len, len(col)) + 2
    # set the column length
    worksheet.set_column(i, i, column_len)
  # save result
  writer.save()


def main():
  while True:
    try:
      scraper()
      sleep(3600*3)
    except KeyboardInterrupt:
      exit(1)


if __name__ == "__main__":
  main()
