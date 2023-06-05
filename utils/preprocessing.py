import pandas as pd
import numpy as np


def prepare_df(file):
    # skipfooter -> skip last n rows
    df = pd.read_excel(file, skipfooter=1)
    df = df.replace("< None >", np.NaN)
    # client relevant data
    df = df.dropna()
    # transform to real datetime object
    df["Entry Date"] = df["Entry Date"].apply(lambda x: pd.to_datetime(x, format="%Y-%m-%d"))
    return df

def clean_name(name):
    new_name = name
    forbidden = ["/", "\\", ":", "*", "\"", "?", "<", ">", "|"]
    for char in name:
        if char in forbidden:
            new_name = new_name.replace(char, "-")
    return new_name
