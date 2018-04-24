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
file_name = "MP Workbook JWB 04-11-2018.xlsx"
mp_file = pd.ExcelFile(file_path + file_name)
generic_factors = mp_file.parse(sheet_name="Generic HV & WD", header=2, \
        usecols="A,B,C,D,E,F,G", \
        names=["idx", "bw", "mw", "pl", "ms", "md", "water"])
generic_factors.set_index('idx', inplace=True)
generic_factors['dcm'] = generic_factors.index.tolist()
generic_factors['step'] = ['generic'] * len(generic_factors)
assignments = mp_file.parse(sheet_name="MP_new", header=None, skiprows=25, \
        usecols="A,B,C,D,F", \
        names=["dca", "area_ac", "area_sqmi", "base", "step0"])
assignments.set_index('dca', inplace=True)
# base case
base_assignments = [generic_factors.index==assignments.base[x] \
        for x in range(0, len(assignments))]
base_assignments = [base_assignments[x].astype(int) for x in \
        range(0, len(base_assignments))]
base = pd.DataFrame(base_assignments, index=assignments.index)
base.columns = generic_factors.index
# step 0
step0_assignments = [generic_factors.index==assignments.step0[x] \
        for x in range(0, len(assignments))]
step0_assignments = [step0_assignments[x].astype(int) for x in \
        range(0, len(step0_assignments))]
step0 = pd.DataFrame(step0_assignments, index=assignments.index)
step0.columns = generic_factors.index

# build up custom habitat and water factor tables
custom_info = mp_file.parse(sheet_name="Custom HV & WD", header=0, \
        usecols="A,B,C,D,E,F,G,H,I", \
        names=["dca", "dcm", "step", "bw", "mw", "pl", "ms", "md", "water"])
custom_filled = custom_info.apply(func.backfill, axis=1, \
        backfill_factors=generic_factors)
custom_factors = {'base': (), 'dwm': (), 'step0': (), 'mp': ()}
for x in ['base', 'dwm', 'step0', 'mp']:
    custom_factors[x] = func.build_custom_steps(x, custom_factors, custom_filled)

# read in constraints (and manual assignments if wanted)
info_file = pd.ExcelFile("/home/john/code/master_proj/DCA-DCM Constraints.xlsx")
# read in manual assignment case (may not be used)
manual = info_file.parse(sheet_name="Manual Scenario", header=2)
# read DCA-DCM constraints file
starting_constraints = info_file.parse(sheet_name="Constraints", header=8)

# read in DCA information
dca_info = pd.read_csv("/home/john/code/master_proj/DCA_top_detailed_MP.csv")
dca_info.drop_duplicates(subset='MP Name', inplace=True)
dca_info['MP Name'] = [x.strip() for x in dca_info['MP Name']]
base_dca_order = [y for x, y in enumerate(base.index.tolist())]
dca_info['sort_index'] = [base_dca_order.index(x) for x in dca_info['MP Name']]
dca_info.sort_values('sort_index', inplace=True)
dca_info.set_index('MP Name', inplace=True)
dca_info = dca_info.loc[:, ['MP Acres', 'Cent_N', 'Cent_E']]
dca_info['sqmi'] = dca_info['MP Acres'] * 0.0015625

# define "soft" transition DCMs?
soft_dcms = ['Tillage', 'Sand Fences']
soft_idx = [x for x, y in enumerate(generic_factors.index.tolist()) if y in soft_dcms]

# set limits and toggles - MAKE CHANGES HERE
allow_sand_fences = False
if not allow_sand_fences:
    starting_constraints.loc[:, 'Sand Fences'] = 0
hard_limit = 3
soft_limit = 4.5
habitat_minimum = 0.9
use_custom_factors = False

# format data for analysis
base_case = np.array(base).copy()
starting_case = np.array(step0).copy()

# evaluate base case habitat and water usage
base_assignments = func.get_assignments(base_case, base.index.tolist(), \
        generic_factors.index.tolist())
if use_custom_factors:
    base_factors = func.build_factor_table(base_assignments, custom_factors, \
            generic_factors, 'step0')
else:
    base_factors = func.build_factor_table(base_assignments, custom_factors, \
            generic_factors, 'generic')
base_total = base_factors.loc[:, ['bw', 'mw', 'pl', 'ms', 'md', 'water']]\
        .multiply(dca_info['MP Acres'], axis=0).sum()

# evaluate starting case compared to base for initial priority
start_assignments = func.get_assignments(starting_case, base.index.tolist(), \
        generic_factors.index.tolist())
if use_custom_factors:
    start_factors = func.build_factor_table(start_assignments, custom_factors, \
            generic_factors, 'step0')
else:
    start_factors = func.build_factor_table(start_assignments, custom_factors, \
            generic_factors, 'generic')
start_total = start_factors.loc[:, ['bw', 'mw', 'pl', 'ms', 'md', 'water']]\
        .multiply(dca_info['MP Acres'], axis=0).sum()

starting_percent = start_total/base_total

# initialize ariables before loop - DO NOT MAKE CHANGES HERE
tick = 0
new_constraints = np.array(starting_constraints).copy()
new_case = starting_case.copy()
new_percent = starting_percent.copy()
new_total = starting_total.copy()
step_info = {}
hard_transition = 0
soft_transition = 0
tracking = []
# set priorities for initial change
priority = func.prioritize(starting_percent, habitat_minimum)

for step in range(0, 6):
    print 'step ' + str(step)
    step_info[step] = {'totals': new_total, 'values': new_percent, \
            'hard transition': hard_transition, \
            'soft_transition': soft_transition, 'changes': tracking, \
            'assignments': new_case}
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
        new_assignments = func.get_assignments(eval_case, base.index.tolist(), \
                generic_factors.index.tolist())
        if use_custom_factors:
            factors = func.build_factor_table(new_assignments, custom_factors, \
                generic_factors, 'mp')
        else:
            factors = func.build_factor_table(new_assignments, custom_factors, \
                generic_factors, 'generic')
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
                    eval_case[dca], factors, priority)
                if case_eval['smart']:
                    dca_assigns.append(allowed_cases[dca][case])
                    benefit1.append(case_eval['benefit1'])
                    benefit2.append(case_eval['benefit2'])
            if not benefit1:
                best_assigns = (0, 0, eval_case[dca].tolist(), dca)
            else:
                best_assigns = [(x, y, z, dca) for x,y,z \
                        in sorted(zip(benefit1, benefit2, dca_assigns), \
                            key=lambda x: (x[0], x[1]), reverse=True)][0]
            smart_cases.append([best_assigns, (0, 0, eval_case[dca].tolist(), dca)])
        n_smart = [len(x) for x in smart_cases]
        smartest = sorted(smart_cases, key=lambda x: (x[0], x[1]), reverse=True)

        try:
            nn = 0
            if smartest[nn][0][2].index(1) in soft_idx:
                while soft_transition + \
                        dca_info.iloc[smartest[nn][0][3]]['sqmi'] > soft_limit:
                    nn += 1
                soft_transition += dca_info.iloc[smartest[nn][0][3]]['sqmi']
            else:
                while hard_transition + \
                        dca_info.iloc[smartest[nn][0][3]]['sqmi'] > hard_limit:
                    nn += 1
                hard_transition += dca_info.iloc[smartest[nn][0][3]]['sqmi']
        except:
            break
        if smartest[nn][0][0] == 0:
            break

        tracking.append({'dca': dca_info.index.tolist()[smartest[nn][0][3]],
                'from': factors.index.tolist()[\
                        previous_case[smartest[nn][0][3]].tolist().index(1)],
                'to': factors.index.tolist()[smartest[nn][0][2].index(1)]})
        new_case = np.zeros(eval_case.shape)
        for idx, item in enumerate(eval_case):
            if idx ==smartest[nn][0][3]:
                new_case[idx] = np.array(smartest[nn][0][2])
            else:
                new_case[idx] = item
        new_constraints = np.zeros(constraints.shape)
        for idx, item in enumerate(constraints):
            if idx ==smartest[nn][0][3]:
                new_constraints[idx] = np.array(smartest[nn][0][2])
            else:
                new_constraints[idx] = item
        new_total = func.evaluate_case(new_case, factors, dca_info['MP Acres']).sum()
        new_percent = new_total/base_total
        priority = func.prioritize(new_percent, habitat_minimum)

total_water_savings = starting_total['water'] - new_total['water']
print 'Finished!'
print 'Total Water Savings = ' + str(total_water_savings) + ' acre-feet/year'
if use_custom_habitat:
    print "Used custom habitat values for Base Case."
else:
    print "Used standard habitat values (no custom habitat values for Base \
    Case)."
if allow_sand_fences:
    print "Use of Sand Fences for dust control was allowed."
else:
    print "Use of Sand Fences for dust control was not allowed."
