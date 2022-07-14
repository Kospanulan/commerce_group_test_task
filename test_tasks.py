import pandas as pd
from datetime import datetime, timedelta
from pyexcelerate import Workbook
from pymongo import MongoClient
import numpy as np

pd.options.mode.chained_assignment = None


def to_excel(df, wb, sheet_name):
    file_name = 'results.xlsx'

    res = [df.columns] + list(df.values)
    wb.new_sheet(sheet_name, data=res)
    wb.save(file_name)


def to_mongo(df, db, collection_name):
    data_dict = df.to_dict("records")
    collection1 = db[collection_name]
    collection1.insert_many(data_dict)


def task1(df, wb, db):
    for_18btw21 = datetime.strptime('2022-01-01T09:00:00', "%Y-%m-%dt%H:%M:%S")

    df.loc[(df['Age'] >= 18) &
           (df['Age'] <= 21) &
           (df['Job'].str.contains('Developer')), 'TimeToEnter'] = for_18btw21

    df.loc[(df['Age'] > 21) &
           (df['Job'].str.contains('Developer')), 'TimeToEnter'] = for_18btw21 + timedelta(minutes=15)

    df.dropna(axis=0, inplace=True)
    df.reset_index(inplace=True, drop=True)

    sheet_name = 'sheet_task1'
    to_excel(df, wb, sheet_name)

    collection_name = "18MoreAnd21andLess"
    to_mongo(df, db, collection_name)


def task2(df, wb, db):
    for_35btw40 = datetime.strptime('2022-01-01T11:00:00', "%Y-%m-%dt%H:%M:%S")

    df.loc[(~df['Job'].str.contains('Developer')) &
           (~df['Job'].str.contains('Manager')) &
           (df['Age'] >= 35) &
           (df['Age'] <= 40), 'TimeToEnter'] = for_35btw40

    df.replace({np.nan: for_35btw40 + timedelta(minutes=30)}, inplace=True)
    print(df)

    sheet_name = 'sheet_task2'
    to_excel(df, wb, sheet_name)

    collection_name = "35AndMore"
    to_mongo(df, db, collection_name)


def task3(df, wb, db):
    for_arch = datetime.strptime('2022-01-01T10:30:00', "%Y-%m-%dt%H:%M:%S")

    df.loc[(df['Job'].str.contains('architect')), 'TimeToEnter'] = for_arch
    df.replace({np.nan: for_arch + timedelta(minutes=10)}, inplace=True)

    print(df)

    sheet_name = 'sheet_task3'
    to_excel(df, wb, sheet_name)

    collection_name = "ArchitectEnterTime"
    to_mongo(df, db, collection_name)


if __name__ == "__main__":
    client = MongoClient("mongodb://localhost:27017/")
    db = client["commerce_group"]

    wb = Workbook()
    format_for_time = "%Y-%m-%dt%H:%M:%S"
    d = np.array(
        [[1, 'Alex', 'Smur', 21, 'Python Developer', datetime.strptime('2022-01-01T09:45:12', format_for_time)],
         [2, 'Justin', 'Forman', 25, 'Java Developer', datetime.strptime('2022-01-01T11:50:25', format_for_time)],
         [3, 'Set', 'Carey', 35, 'Project Manager', datetime.strptime('2022-01-01T10:00:45', format_for_time)],
         [4, 'Carlos', 'Carey', 40, 'Enterprise architect', datetime.strptime('2022-01-01T09:07:36', format_for_time)],
         [5, 'Gareth', 'Chapman', 19, 'Python Developer', datetime.strptime('2022-01-01T11:54:10', format_for_time)],
         [6, 'John', 'James', 27, 'IOS Developer', datetime.strptime('2022-01-01T09:56:40', format_for_time)],
         [7, 'Bob', 'James', 25, 'Python Developer', datetime.strptime('2022-01-01T09:52:45', format_for_time)]])

    df = pd.DataFrame(data=d, columns=['Id', 'Name', 'Surname', 'Age', 'Job', 'Datetime'])
    print(df)

    df1 = df.copy()
    task1(df1, wb, db)

    df2 = df.copy()
    task2(df2, wb, db)

    df3 = df.copy()
    task3(df3, wb, db)
