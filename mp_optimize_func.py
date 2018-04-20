import pandas as pd
from time import time
import numpy as np

def countup(x):
    return sum(1 for a in x)

def evaluate_case(case, factors, areas):
    """With an assignment matrix, calculate total habitat acreage and water
    use for a MP case.
    case = assignment matrix for case being evaluated (DataFrame).
    factors = habitat and water use factors (DataFrame).
    areas = areas of DCAs in acres (Series, in same DCA order as case).
    """
    case_factors = pd.DataFrame(np.empty([len(case), len(factors.columns)]), \
            index=areas.index, columns=factors.columns.tolist())
    for x in range(0, len(factors.columns)):
        case_factors.iloc[:, x] = case.dot(factors.iloc[:, x]) * areas
    return case_factors

def evaluate_dca_change(case, previous_case, factors, priority):
    previous_case_factors = factors.iloc[previous_case.tolist().index(1)]
    case_factors = factors.iloc[case.index(1)]
    if priority[1]=='water':
        smart = case_factors['water'] - previous_case_factors['water'] <= 0
        benefit1 = previous_case_factors['water'] - case_factors['water']
        benefit2 = case_factors[priority[2]] - previous_case_factors[priority[2]]
    else:
        smart = case_factors[priority[1]] - previous_case_factors[priority[1]] >= 0
        benefit1 = case_factors[priority[1]] - previous_case_factors[priority[1]]
        benefit2 = previous_case_factors['water'] - case_factors['water']
    return {'smart': smart, 'benefit1':benefit1, 'benefit2':benefit2}


def single_factor_total(case, dcm_factors, dca_areas):
    """With an assignment matrix for a MP scenario, calculate total acreage
    (or acre-feet/year) for a single guild habitat (or water usage).
    case = assignment matrix for scenario being evaluated (array or DataFrame).
    """
    area_values = case.dot(dcm_factors) * dca_areas
    return area_values.sum()

# constraint functions
def compare_value(case, factors, areas, check_case, percent):
    """
    Check whether a factor value has decrease from a previous scenario.
    """
    val = single_factor_total(np.array(case), factors, areas)
    check_val = single_factor_total(np.array(check_case), factors, areas)
    return val < percent * check_val

def transition_area(case, check_case, dcms_list, dca_areas):
    transition = {'hard': [0] * len(case), 'soft': [0] * len(case)}
    soft_dcms = ['Tillage', 'Brine', 'Till-Brine']
    soft_indices = [x for x, y in enumerate(dcms_list) if y in soft_dcms]
    for row in range(0, len(case)):
        change = not all(case[row] == check_case[row])
        if change and case[row].index(1) in soft_indices:
            transition['soft'][row] = 1
        if change and case[row].index(1) not in soft_indices:
            transition['hard'][row] = 1
    hard_sqmi = pd.Series(transition['hard'] * dca_areas).sum()
    soft_sqmi = pd.Series(transition['soft'] * dca_areas).sum()
    return {'hard': hard_sqmi, 'soft': soft_sqmi}

def hard_area_check(case, check_case, soft_indices, dca_areas, hard_limit):
    transition = [0] * len(case)
    for row in range(0, len(case)):
        change = not all(case[row] == check_case[row])
        if change and list(case[row]).index(1) not in soft_indices:
            transition[row] = 1
    return pd.Series(transition * dca_areas).sum() < hard_limit

def benefit_check(case, check_sum, factors, areas):
    case_sum = evaluate_case(case, factors, areas).sum()
    hab_decrease = [case_sum.iloc[i] < check_sum.iloc[i] for i in range(0, 4)]
    water_increase = case_sum.iloc[5] > check_sum.iloc[5]
    return not (all(hab_decrease) and water_increase)

def overload_increase(case, check_sum, hab, factors, areas):
    case_sum = evaluate_case(case, factors, areas).sum()
    hab_increase = case_sum[hab] > check_sum[hab]
    water_increase = case_sum['water_af/y'] > check_sum['water_af/y']
    return not (hab_increase and water_increase)

def prioritize(value_percents):
    if any([x < 0.9 for x in value_percents[0:5]]):
        return {1: value_percents.idxmin(), 2: 'water'}
    else:
        return {1: 'water', 2: value_percents.idxmin()}
