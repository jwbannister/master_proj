import pandas as pd
from time import time
import numpy as np
import itertools as it

def evaluate_dca_change(dca_case, previous_case, previous_factors, custom_factors, \
        priority, dca_idx, dca_info):
    previous_case_factors = previous_factors.loc[previous_case.name]
    dca_name = dca_info.iloc[dca_idx].name
    dcm_name = custom_factors['dcm'].index.tolist()[dca_case.index(1)]
    if (dca_name, dcm_name) in [(x, y) for x, y in custom_factors['mp'].index.tolist()]:
       case_factors = custom_factors['mp'].loc[dca_name, dcm_name]
    else:
       case_factors = custom_factors['dcm'].iloc[dca_case.index(1)]
    if priority[1]=='water':
        smart = case_factors['water'] - previous_case_factors['water'] < 0
        benefit1 = previous_case_factors['water'] - case_factors['water']
        benefit2 = case_factors[priority[2]] - previous_case_factors[priority[2]]
    else:
        smart = case_factors[priority[1]] - previous_case_factors[priority[1]] > 0
        benefit1 = case_factors[priority[1]] - previous_case_factors[priority[1]]
        benefit2 = previous_case_factors['water'] - case_factors['water']
    return {'smart': smart, 'benefit1':benefit1, 'benefit2':benefit2}

def prioritize(value_percents, minimum_hab):
    if any([x < minimum_hab for x in value_percents[0:5]]):
        return {1: value_percents[0:5].idxmin(), 2: 'water'}
    else:
        return {1: 'water', 2: value_percents[0:5].idxmin()}

def backfill(row, backfill_data, columns_list):
    """
    row = Series with index of columns_list and name (X, Y) where Y = backfill_factors.index
    backfill_factors = DataFrame with all columns_list columns
    columns_list = list
    """
    for col in columns_list:
        if np.isnan(row[col]):
            row[col] = backfill_data.loc[row.name[1], col]
    return row

def build_custom_steps(step, step_list, data):
    """
    step = string
    step_list = list
    data = DataFrame with column 'step'
    """
    factors = data.loc[data['step']==step_list[0], :].copy()
    sub_steps = [step_list[x] for x in range(1, step_list.index(step)+1)]
    for sub in sub_steps:
        sub_df = data.loc[data['step']==sub, :].copy()
        for idx in factors.index:
            if idx in sub_df.index:
                factors.drop(idx, inplace=True)
        factors = factors.append(sub_df)
    factors.drop('step', axis=1, inplace=True)
    return factors

def get_assignments(case, dca_list, dcm_list):
    assignments = pd.DataFrame([dcm_list[row.tolist().index(1)] \
            for index, row in case.iterrows()], index=dca_list, columns=['dcm'])
    return assignments

def build_case_factors(lake_case, custom_factors, stp):
    factors = pd.DataFrame()
    for idx in range(0, len(lake_case)):
        dca_name = lake_case.iloc[idx].name
        dca_case = lake_case.iloc[idx]
        dcm_name = dca_case[dca_case==1].index[0]
        dca_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dca')) \
                if y==dca_name]
        dcm_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dcm')) \
                if y==dcm_name]
        custom_idx = [x for x in dca_idx if x in dcm_idx]
        if len(custom_idx)>0:
            tmp = custom_factors[stp].iloc[custom_idx].copy()
            tmp['dca'] = lake_case.index.tolist()[idx]
            factors = factors.append(tmp)
        else:
            tmp = custom_factors['dcm'].loc[dcm_name].copy()
            tmp['dca'] = lake_case.index.tolist()[idx]
            factors = factors.append(tmp)
    factors.set_index('dca', inplace=True)
    return factors

def calc_totals(case, custom_factors, step, dca_info):
    dca_list = dca_info.index.tolist()
    factors = build_case_factors(case, custom_factors, step)
    return factors.multiply(np.array(dca_info['area_ac']), axis=0).sum()

