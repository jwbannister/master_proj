import pandas as pd
from time import time
import numpy as np

def evaluate_dca_change(case, previous_case, previous_factors, custom_factors, \
        priority, dca_idx, dca_info):
    previous_case_factors = previous_factors.loc[previous_case.name]
    dca_name = dca_info.iloc[dca_idx].name
    dcm_name = custom_factors['generic'].index.tolist()[case.index(1)]
    if (dca_name, dcm_name) in [(x, y) for x, y in custom_factors['mp'].index.tolist()]:
       case_factors = custom_factors['mp'].loc[dca_name, dcm_name]
    else:
       case_factors = custom_factors['generic'].iloc[case.index(1)]
    if priority[1]=='water':
        smart = case_factors['water'] - previous_case_factors['water'] <= 0
        benefit1 = previous_case_factors['water'] - case_factors['water']
        benefit2 = case_factors[priority[2]] - previous_case_factors[priority[2]]
    else:
        smart = case_factors[priority[1]] - previous_case_factors[priority[1]] >= 0
        benefit1 = case_factors[priority[1]] - previous_case_factors[priority[1]]
        benefit2 = previous_case_factors['water'] - case_factors['water']
    return {'smart': smart, 'benefit1':benefit1, 'benefit2':benefit2}

def prioritize(value_percents, minimum_hab):
    if any([x < minimum_hab for x in value_percents[0:5]]):
        return {1: value_percents[0:5].idxmin(), 2: 'water'}
    else:
        return {1: 'water', 2: value_percents[0:5].idxmin()}

def backfill(row, backfill_factors):
    for col in ['bw', 'mw', 'pl', 'ms', 'md', 'water']:
        if np.isnan(row[col]):
            row[col] = backfill_factors.loc[row['dcm'], col]
    return row

def build_custom_steps(stp, custom_factors, custom_filled):
    x_walk = {'base': 'Base', 'dwm': 'DWM', 'step0': 'Zero', 'mp': 'MP'}
    previous_walk = {'base':'', 'dwm': 'base', 'step0': 'dwm', 'mp': 'step0'}
    prev = previous_walk[stp]
    excel_stp = x_walk[stp]
    if stp == 'base':
        custom_factors['base'] = custom_filled.loc[custom_filled['step']=='Base', :]
    else:
        tmp = custom_filled.loc[custom_filled['step']==excel_stp, :]
        carry = [x not in tmp['dca'].tolist() for \
                x in custom_factors[prev]['dca'].tolist()]
        custom_factors[stp] = tmp.append(custom_factors[prev].iloc[carry, :])
    return custom_factors[stp]

def get_assignments(case, dca_list, dcm_list):
    assignments = pd.DataFrame([dcm_list[row.tolist().index(1)] \
            for index, row in case.iterrows()], index=dca_list, columns=['dcm'])
    return assignments

def build_factor_table(assignments, custom_factors, stp):
    factors = pd.DataFrame()
    for idx in range(0, len(assignments)):
        dcm_spot = assignments.iloc[idx].tolist().index(1)
        dcm_name = custom_factors['generic'].index.tolist()[dcm_spot]
        dca_name = assignments.index.tolist()[idx]
        dca_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dca')) \
                if y==dca_name]
        dcm_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dcm')) \
                if y==dcm_name]
        custom_idx = [x for x in dca_idx if x in dcm_idx]
        if len(custom_idx)>0:
            tmp = custom_factors[stp].iloc[custom_idx].copy()
            tmp['dca'] = assignments.index.tolist()[idx]
            factors = factors.append(tmp)
        else:
            tmp = custom_factors['generic'].loc[dcm_name].copy()
            tmp['dca'] = assignments.index.tolist()[idx]
            factors = factors.append(tmp)
    factors.set_index('dca', inplace=True)
    return factors

def calc_totals(case, custom_factors, step, dca_info):
    dca_list = dca_info.index.tolist()
    factors = build_factor_table(case, custom_factors, step)
    return factors.multiply(np.array(dca_info['area_ac']), axis=0).sum()

