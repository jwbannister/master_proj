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
file_name = "MP Workbook JWB 04-25-2018.xlsx"
mp_file = pd.ExcelFile(file_path + file_name)
generic_factors = mp_file.parse(sheet_name="Generic HV & WD", header=2, \
        usecols="A,B,C,D,E,F,G", \
        names=["dcm", "bw", "mw", "pl", "ms", "md", "water"])
generic_factors['step'] = ['generic'] * len(generic_factors)
generic_factors.set_index('dcm', drop=False, inplace=True)
assignments = mp_file.parse(sheet_name="MP_new", header=None, skiprows=25, \
        usecols="A,B,C,D,F", \
        names=["dca", "area_ac", "area_sqmi", "base", "step0"])
assignments.set_index('dca', inplace=True)
# read in dca info
dca_info = assignments.drop(['base', 'step0'], axis=1)
# base case
base_assignments = [generic_factors['dcm']==assignments.base[x] \
        for x in range(0, len(assignments))]
base_assignments = [base_assignments[x].astype(int) for x in \
        range(0, len(base_assignments))]
base = pd.DataFrame(base_assignments, index=assignments.index)
base.columns = generic_factors['dcm']
# step 0
step0_assignments = [generic_factors['dcm']==assignments.step0[x] \
        for x in range(0, len(assignments))]
step0_assignments = [step0_assignments[x].astype(int) for x in \
        range(0, len(step0_assignments))]
step0 = pd.DataFrame(step0_assignments, index=assignments.index)
step0.columns = generic_factors['dcm']
# read DCA-DCM constraints file
start_constraints = mp_file.parse(sheet_name="Constraints", header=8, \
        usecols="A:AF")

# build up custom habitat and water factor tables
custom_info = mp_file.parse(sheet_name="Custom HV & WD", header=0, \
        usecols="A,B,C,D,E,F,G,H,I", \
        names=["dca", "dcm", "step", "bw", "mw", "pl", "ms", "md", "water"])
custom_filled = custom_info.apply(func.backfill, axis=1, \
        backfill_factors=generic_factors)
custom_factors = {'base': (), 'dwm': (), 'step0': (), 'mp': ()}
for x in ['base', 'dwm', 'step0', 'mp']:
    custom_factors[x] = func.build_custom_steps(x, custom_factors, custom_filled)
for x in ['base', 'dwm', 'step0', 'mp']:
    custom_factors[x].set_index([custom_factors[x]['dca'], \
            custom_factors[x]['dcm']], inplace=True)
    custom_factors[x].drop(['dca', 'dcm', 'step'], axis=1, inplace=True)
# format generic to similar structure to custom
generic_factors.drop(['dcm', 'step'], axis=1, inplace=True)

# define "soft" transition DCMs
soft_dcms = ['Tillage', 'Brine', 'Till-Brine', 'Sand Fences']
soft_idx = [x for x, y in enumerate(generic_factors.index.tolist()) if y in soft_dcms]

# set limits and toggles - MAKE CHANGES HERE
allow_sand_fences = True
if not allow_sand_fences:
    start_constraints.loc[:, 'Sand Fences'] = 0
hard_limit = 3
soft_limit = 3
habitat_minimum = 0.9 + 0.01
use_custom_factors = True

# format data for analysis
base_case = base.copy()
start_case = step0.copy()

base_total = func.calc_totals(base_case, custom_factors, generic_factors, \
        'base', use_custom_factors, dca_info)
start_total = func.calc_totals(start_case, custom_factors, generic_factors, \
        'step0', use_custom_factors, dca_info)
start_percent = start_total/base_total

# initialize ariables before loop - DO NOT MAKE CHANGES HERE
new_constraints = np.array(start_constraints).copy()
new_case = start_case.copy()
new_assignments = func.get_assignments(new_case, base.index.tolist(), \
        generic_factors.index.tolist())
if use_custom_factors:
    factors = func.build_factor_table(new_assignments, custom_factors, \
        generic_factors, 'mp')
else:
    factors = func.build_factor_table(new_assignments, custom_factors, \
        generic_factors, 'generic')
new_percent = start_percent.copy()
new_total = start_total.copy()
step_info = {}
hard_transition = 0
soft_transition = 0
tracking = []
# set priorities for initial change
priority = func.prioritize(start_percent, habitat_minimum)
step_info[0] = {'totals': new_total, 'percent_base': new_percent, \
        'hard transition': hard_transition, \
        'soft_transition': soft_transition, 'changes': tracking, \
        'assignments': new_assignments}

for step in range(1, 6):
    print 'step ' + str(step)
    hard_transition = 0
    soft_transition = 0
    tracking = []
    while hard_transition < hard_limit and soft_transition < soft_limit:
        print "hard = " + str(round(hard_transition, 2)) + ", soft = " +\
                str(round(soft_transition, 2)) + ", " + str(priority[1]) + " = " +\
                str(round(new_percent[priority[1]], 3)) + ", " + str(priority[2]) +\
                " = " + str(round(new_percent[priority[2]], 3))
        constraints = new_constraints.copy()
        eval_case = new_case.copy()
        allowed_cases = []
        for j in range(0, len(constraints)):
            tmp = constraints[j].tolist()
            tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
            a = []
            for d in tmp_ind:
                b = [0 for x in tmp]
                b[d] = 1
                a.append(b)
            allowed_cases.append(a)
        n_allowed = [len(x) for x in allowed_cases]

        smart_cases = []
        for dca in range(0, len(allowed_cases)):
            benefit1 = []
            benefit2 = []
            dca_assigns = []
            for case in range(0, len(allowed_cases[dca])):
                case_eval = func.evaluate_dca_change(allowed_cases[dca][case], \
                    eval_case.iloc[dca], factors, generic_factors, priority)
                if case_eval['smart']:
                    dca_assigns.append(allowed_cases[dca][case])
                    benefit1.append(case_eval['benefit1'])
                    benefit2.append(case_eval['benefit2'])
            if not benefit1:
                best_assigns = (0, 0, eval_case.iloc[dca].tolist(), dca)
            else:
                best_assigns = [(x, y, z, dca) for x,y,z \
                        in sorted(zip(benefit1, benefit2, dca_assigns), \
                            key=lambda x: (x[0], x[1]), reverse=True)][0]
            smart_cases.append([best_assigns, (0, 0, eval_case.iloc[dca].tolist(), dca)])
        n_smart = [len(x) for x in smart_cases]
        smartest = sorted(smart_cases, key=lambda x: (x[0], x[1]), reverse=True)

        try:
            nn = 0
            if smartest[nn][0][2].index(1) in soft_idx:
                while soft_transition + \
                        dca_info.iloc[smartest[nn][0][3]]['area_sqmi'] > soft_limit:
                    nn += 1
                soft_transition += dca_info.iloc[smartest[nn][0][3]]['area_sqmi']
            else:
                while hard_transition + \
                        dca_info.iloc[smartest[nn][0][3]]['area_sqmi'] > hard_limit:
                    nn += 1
                hard_transition += dca_info.iloc[smartest[nn][0][3]]['area_sqmi']
        except:
            break
        if smartest[nn][0][0] == 0:
            break

        tracking.append({'dca': dca_info.index.tolist()[smartest[nn][0][3]],
                'from': eval_case.columns.tolist()[\
                        eval_case.iloc[smartest[nn][0][3]].tolist().index(1)],
                'to': eval_case.columns.tolist()[smartest[nn][0][2].index(1)]})
        new_case = eval_case.copy()
        new_case.iloc[smartest[nn][0][3]] = np.array(smartest[nn][0][2])
        new_constraints = constraints.copy()
        new_constraints[smartest[nn][0][3]] = np.array(smartest[nn][0][2])
        new_assignments = func.get_assignments(new_case, base.index.tolist(), \
                generic_factors.index.tolist())
        if use_custom_factors:
            factors = func.build_factor_table(new_assignments, custom_factors, \
                generic_factors, 'mp')
        else:
            factors = func.build_factor_table(new_assignments, custom_factors, \
                generic_factors, 'generic')
        new_total = factors.multiply(dca_info['area_ac'], axis=0).sum()
        new_percent = new_total/base_total
        priority = func.prioritize(new_percent, habitat_minimum)
    step_info[step] = {'totals': new_total, 'percent_base': new_percent, \
            'hard transition': hard_transition, \
            'soft_transition': soft_transition, 'changes': tracking, \
            'assignments': new_assignments}

total_water_savings = base_total['water'] - new_total['water']
print 'Finished!'
print 'Total Water Savings = ' + str(total_water_savings) + ' acre-feet/year'
if use_custom_factors:
    print "Used custom habitat values for Base Case."
else:
    print "Used standard habitat values (no custom habitat values for Base \
    Case)."
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

