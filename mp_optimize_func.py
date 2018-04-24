import pandas as pd
from time import time
import numpy as np

def evaluate_dca_change(case, previous_case, factors, generic_factors, priority):
    previous_case_factors = factors.loc[previous_case.name]
    case_factors = generic_factors.iloc[case.index(1)]
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

def build_factor_table(assignments, custom_factors, generic_factors, stp):
    factors = pd.DataFrame()
    for idx in assignments.index.tolist():
        dcm = assignments.loc[idx]['dcm']
        if stp == 'generic':
            tmp = generic_factors.loc[dcm].copy()
            tmp['dca'] = idx
            factors = factors.append(tmp)
        else:
            dca_idx = [x for x, y in \
                    enumerate(custom_factors[stp].index.get_level_values('dca')) \
                    if y==idx]
            dcm_idx = [x for x, y in \
                    enumerate(custom_factors[stp].index.get_level_values('dcm')) \
                    if y==dcm]
            custom_idx = [x for x in dca_idx if x in dcm_idx]
            if len(custom_idx)>0:
                tmp = custom_factors[stp].iloc[custom_idx].copy()
                tmp['dca'] = idx
                factors = factors.append(tmp)
            else:
                tmp = generic_factors.loc[dcm].copy()
                tmp['dca'] = idx
                factors = factors.append(tmp)
    factors.set_index('dca', inplace=True)
    return factors

def calc_totals(case, custom_factors, generic_factors, step, \
        use_custom_factors, dca_info):
    dca_list = dca_info.index.tolist()
    dcm_list = generic_factors.index.tolist()
    assignments = get_assignments(case, dca_list, dcm_list)
    if use_custom_factors:
        factors = build_factor_table(assignments, custom_factors, \
                generic_factors, step)
    else:
        factors = build_factor_table(assignments, custom_factors, \
                generic_factors, 'generic')
    return factors.multiply(dca_info['area_ac'], axis=0).sum()

