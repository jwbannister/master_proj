import pandas as pd
import numpy as np
import itertools as it
import csv
from multiprocessing import Pool
import mp_optimize_func as func
from time import time
from collections import Counter

# read data from original Master Project planning workbook
file_path = "/home/john/airsci/owens/Master Project & Cost Benefit/"
file_name = "MP Workbook JWB 04-27-2018.xlsx"
mp_file = pd.ExcelFile(file_path + file_name)

# generate habitat and water duty factor tables
dcm_factors = mp_file.parse(sheet_name="Generic HV & WD", header=2, \
        usecols="A,C,D,E,F,G,H", \
        names=["dcm", "bw", "mw", "pl", "ms", "md", "water"])
dcm_factors.set_index('dcm', inplace=True)
# build up custom habitat and water factor tables
custom_info = mp_file.parse(sheet_name="Custom HV & WD", header=0, \
        usecols="A,B,C,D,E,F,G,H,I", \
        names=["dca", "dcm", "step", "bw", "mw", "pl", "ms", "md", "water"])
custom_info.set_index(['dca', 'dcm'], inplace=True)
custom_filled = custom_info.apply(func.backfill, axis=1, \
        backfill_data=dcm_factors, \
        columns_list=['bw', 'mw', 'pl', 'ms', 'md', 'water'])
custom_steps = ['base', 'dwm', 'step0', 'mp']
factors = {x: func.build_custom_steps(x, custom_steps, custom_filled) \
        for x in custom_steps}
factors['dcm'] = dcm_factors.copy()

dca_info = mp_file.parse(sheet_name="MP_new", header=None, skiprows=25, \
        usecols="A,B,C,D,F", \
        names=["dca", "area_ac", "area_sqmi", "base", "step0"])
dca_info.set_index('dca', inplace=True)

# known cases
lake_case = {'base': [], 'step0': []}
for case in lake_case.keys():
    assignments = [np.array(factors['dcm'].index.get_level_values('dcm')) \
            == dca_info[case][x] for x in range(0, len(dca_info))]
    assignments = [assignments[x].astype(int).tolist() \
            for x in range(0, len(assignments))]
    case_df = pd.DataFrame(assignments)
    case_df.index = dca_info.index
    case_df.columns = factors['dcm'].index.tolist()
    lake_case[case] = case_df

# read DCA-DCM constraints file
start_constraints = mp_file.parse(sheet_name="Constraints", header=8, \
        usecols="A:AF")

# define "soft" transition DCMs
soft_dcms = ['Tillage', 'Brine', 'Till-Brine', 'Sand Fences']
soft_idx = [x for x, y in enumerate(factors['dcm'].index.tolist()) if y in soft_dcms]

# set limits and toggles - MAKE CHANGES HERE
allow_sand_fences = True
if not allow_sand_fences:
    start_constraints.loc[:, 'Sand Fences'] = 0
hard_limit = 3
soft_limit = 3
habitat_minimum = 0.9 + 0.01

total = {}
for case in lake_case.keys():
    total[case] = func.calc_totals(lake_case[case], factors, case, dca_info)

# initialize ariables before loop - DO NOT MAKE CHANGES HERE
dcm_list = factors['dcm'].index.tolist()
dca_list = lake_case['base'].index.get_level_values('dca').tolist()
new_constraints = start_constraints.copy()
new_case = lake_case['step0'].copy()
new_assignments = func.get_assignments(new_case, dca_list, dcm_list)
case_factors = func.build_case_factors(new_case, factors, 'mp')
new_percent = total['step0']/total['base']
new_total = total['step0'].copy()
step_info = {}
tracking = []
hard_transition = 0
soft_transition = 0
# set priorities for initial change
priority = func.prioritize(new_percent, habitat_minimum)
step_info[0] = {'totals': new_total, 'percent_base': new_percent, \
        'hard transition': hard_transition, \
        'soft_transition': soft_transition, 'changes': tracking, \
        'assignments': new_assignments}
for step in range(1, 6):
    print 'step ' + str(step)
    hard_transition = 0
    soft_transition = 0
    tracking = []
    while hard_transition < hard_limit or soft_transition < soft_limit:
        print "hard = " + str(round(hard_transition, 2)) + ", soft = " +\
                str(round(soft_transition, 2)) + ", " + str(priority[1]) + " = " +\
                str(round(new_percent[priority[1]], 3)) + ", " + str(priority[2]) +\
                " = " + str(round(new_percent[priority[2]], 3))
        constraints = new_constraints.copy()
        eval_case = new_case.copy()

        allowed_cases = []
        for dca in range(0, len(constraints)):
            tmp = constraints.iloc[dca].tolist()
            tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
            a = []
            for dcm in tmp_ind:
                b = [0 for x in tmp]
                b[dcm] = 1
                a.append(b)
            allowed_cases.append(a)
        n_allowed = [len(x) for x in allowed_cases]

        smart_cases = {'soft': [], 'hard': []}
        smartest = {'soft': [], 'hard': []}
        for dca in range(0, len(allowed_cases)):
            benefit1 = {'soft': [], 'hard': []}
            benefit2 = {'soft': [], 'hard': []}
            dca_assigns = {'soft': [], 'hard': []}
            for case in range(0, len(allowed_cases[dca])):
                case_eval = func.evaluate_dca_change(allowed_cases[dca][case], \
                    eval_case.iloc[dca], case_factors, factors, \
                    priority, dca, dca_info)
                if allowed_cases[dca][case].index(1) in soft_idx:
                    flag = 'soft'
                else:
                    flag='hard'
                if case_eval['smart']:
                    dca_assigns[flag].append(allowed_cases[dca][case])
                    benefit1[flag].append(case_eval['benefit1'])
                    benefit2[flag].append(case_eval['benefit2'])
            for flag in ['soft', 'hard']:
                no_change = (0, 0, eval_case.iloc[dca].tolist(), dca, flag)
                if not benefit1[flag]:
                    best = no_change
                else:
                    best = [(x, y, z, dca, flag) for x,y,z \
                            in sorted(zip(benefit1[flag], benefit2[flag], \
                            dca_assigns[flag]), key=lambda x: (x[0], x[1]), \
                            reverse=True)][0]
                smart_cases[flag].append(best)
        n_smart = [len(smart_cases[x]) for x in smart_cases]
        for flag in ['soft', 'hard']:
            smartest[flag] = sorted(smart_cases[flag], key=lambda x: (x[0], x[1]), \
                    reverse=True)

        soft_nn = 0
        hard_nn = 0
        best_change = sorted([smartest['soft'][soft_nn], \
                smartest['hard'][hard_nn]], key=lambda x: (x[0], x[1]), \
                reverse=True)[0]
        try:
            while True:
                best_change = sorted([smartest['soft'][soft_nn], \
                        smartest['hard'][hard_nn]], key=lambda x: (x[0], x[1]), \
                        reverse=True)[0]
                if best_change[4] == 'soft':
                    if soft_transition + \
                        dca_info.iloc[best_change[3]]['area_sqmi'] > soft_limit:
                        soft_nn += 1
                        continue
                    else:
                        soft_transition += dca_info.iloc[best_change[3]]['area_sqmi']
                        break
                else:
                    if hard_transition + \
                        dca_info.iloc[best_change[3]]['area_sqmi'] > hard_limit:
                        hard_nn += 1
                        continue
                    else:
                        hard_transition += dca_info.iloc[best_change[3]]['area_sqmi']
                        break
        except:
            hard_transition = hard_limit + 1
            soft_transition = soft_limit + 1

        if hard_transition < hard_limit or soft_transition < soft_limit:
            tracking.append({'dca': dca_info.index.tolist()[best_change[3]],
                    'from': eval_case.columns.tolist()[\
                            eval_case.iloc[best_change[3]].tolist().index(1)],
                    'to': eval_case.columns.tolist()[best_change[2].index(1)]})
        new_case = eval_case.copy()
        new_case.iloc[best_change[3]] = np.array(best_change[2])
        new_constraints = constraints.copy()
        new_constraints.iloc[best_change[3]] = np.array(best_change[2])
        new_assignments = func.get_assignments(new_case, dca_list, dcm_list)
        case_factors = func.build_case_factors(new_case, factors, 'mp')
        new_total = case_factors.multiply(dca_info['area_ac'], axis=0).sum()
        new_percent = new_total/total['base']
        priority = func.prioritize(new_percent, habitat_minimum)
    step_info[step] = {'totals': new_total, 'percent_base': new_percent, \
            'hard transition': hard_transition, \
            'soft_transition': soft_transition, 'changes': tracking, \
            'assignments': new_assignments}
total_water_savings = total['base']['water'] - new_total['water']
print 'Finished!'
print 'Total Water Savings = ' + str(total_water_savings) + ' acre-feet/year'
if allow_sand_fences:
    print "Use of Sand Fences for dust control was allowed."
else:
    print "Use of Sand Fences for dust control was not allowed."
assignment_output = new_assignments.copy()
assignment_output['step'] = 0
for i in range(1, 6):
    changes = [x['dca'] for x in step_info[i]['changes']]
    flag = [x in changes for x in assignment_output.index.tolist()]
    assignment_output.loc[flag, 'step'] = i
assignment_output.to_csv("/home/john/Desktop/mp_optimal_assignments.csv")

