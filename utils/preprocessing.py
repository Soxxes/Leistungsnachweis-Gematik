import sys
import logging

import pandas as pd
import numpy as np

from utils.utils import add_logging


@add_logging
def prepare_df(file):
    # skipfooter -> skip last n rows
    df = pd.read_excel(file, skipfooter=1)
    df = df.replace("< None >", np.NaN)
    df["Task Name"].replace(r'^\s*$', np.NaN, regex=True, inplace=True)
    # client relevant data
    df = df.dropna(subset=["Client Name"])
    # transform to real datetime object
    df["Entry Date"] = df["Entry Date"].apply(lambda x: pd.to_datetime(x, format="%Y-%m-%d"))

    assert len(df) != 0, "created DataFrame is empty"
    logging.info("Successfully prepared dataframe.")
    return df


def clean_name(name):
    new_name = name
    forbidden = ["/", "\\", ":", "*", "\"", "?", "<", ">", "|"]
    if isinstance(name, str):
        for char in name:
            if char in forbidden:
                new_name = new_name.replace(char, "-")
    return new_name


@add_logging
def merge_groups(groups, client_info) -> dict:
    merged_groups = {}
    for task_name, task_name_group in groups:
        task_name = task_name.split()[0]
        if task_name in client_info.get("additional_tasks").keys():
            task_name = client_info.get("additional_tasks").get(task_name)
        if merged_groups.get(task_name) is None:
            merged_groups[task_name] = []
        merged_groups[task_name].append(task_name_group)
            
    # groups is list of groups and will be replaced by one merged group for
    # the corresponding mapped task
    for task_name, groups in merged_groups.items():
        task_name_group = pd.concat(groups)
        task_name_group.sort_values("Entry Date", inplace=True)
        merged_groups[task_name] = task_name_group

    logging.info("Successfully merged groups.")
    return merged_groups

