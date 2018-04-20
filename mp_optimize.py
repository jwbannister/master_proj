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
        names=["hdcm", "bw", "mw", "pl", "ms", "md", "water"])
generic_factors.set_index('hdcm', inplace=True)
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

info_file = pd.ExcelFile("/home/john/code/master_proj/DCA-DCM Constraints.xlsx")
# read in manual assignment case (may not be used)
manual = info_file.parse(sheet_name="Manual Scenario", header=2)
# read DCA-DCM constraints file
starting_constraints = info_file.parse(sheet_name="Constraints", header=8)

# read in file with DCA information
dca_info = pd.read_csv("/home/john/code/master_proj/DCA_top_detailed_MP.csv")
dca_info.drop_duplicates(subset='MP Name', inplace=True)
dca_info['MP Name'] = [x.strip() for x in dca_info['MP Name']]
base_dca_order = [y for x, y in enumerate(base.index.tolist())]
dca_info['sort_index'] = [base_dca_order.index(x) for x in dca_info['MP Name']]
dca_info.sort_values('sort_index', inplace=True)
dca_info.set_index('MP Name', inplace=True)

# BUILD TEST DATA
num_dca = 160
num_dcm = 31
base_case = np.array(base.iloc[0:num_dca, 0:num_dcm]).copy()
factors = generic_factors.iloc[0:num_dcm, :].copy().copy()
starting_case = np.array(step0.iloc[0:num_dca, 0:num_dcm]).copy()
# reduce contraints table to match test data
starting_constraints = np.array(starting_constraints.iloc[0:num_dca, 0:num_dcm]).copy()
# reduce DCA info to match test data
dca_info = dca_info.iloc[0:num_dca, 0:num_dcm].copy()
dca_info = dca_info.loc[:, ['MP Acres', 'Cent_N', 'Cent_E']]
dca_info['sqmi'] = dca_info['MP Acres'] * 0.0015625

# which DCMs count as a "soft" transition?
soft_dcms = ['Tillage', 'Brine', 'Till-Brine']
soft_idx = [x for x, y in enumerate(factors.index.tolist()) if y in soft_dcms]

# evaluate base case habitat and water usage
base_total = func.evaluate_case(base_case, factors, dca_info['MP Acres']).sum()
# evaluate starting case compared to base for initial priority
starting_total = func.evaluate_case(starting_case, \
        factors, dca_info['MP Acres']).sum()
starting_percent = starting_total/base_total
priority = func.prioritize(starting_percent)
# initialize variables before loop
hard_transition = 0
soft_transition = 0
tracking = []
tick = 0
new_constraints = starting_constraints.copy()
new_case = starting_case.copy()

while hard_transition < 3:
    print tick
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
        while hard_transition + dca_info.iloc[smartest[nn][0][3]]['sqmi'] > 3:
            nn += 1
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
    priority = func.prioritize(new_percent)
    change_area = dca_info.iloc[smartest[nn][0][3]]['sqmi']
    if smartest[nn][0][2].index(1) in soft_idx:
        soft_transition += change_area
    else:
        hard_transition += change_area
    tick += 1


# which DCAs have a constraint conflict (no smart cases)?
conflicts = [x for x, y in enumerate(n_smart) if y==0]

    # build iterator of all possible assignment combinations
    assignment_combinations = lambda: it.product(*allowed_cases)
    # remove chunks with total hard transition area > 3 sq miles
    filter_hard = lambda: it.ifilter(lambda x: func.hard_area_check\
        (x, previous_case_chunk[i], soft_indices, dca_info_chunk[i]['sqmi'], 3), \
        assignment_combinations())

    # which scenarios have no benefit 
    # (water increase with no habitat increases)
    benefit = \
            it.starmap(func.benefit_check, \
            ((np.array(x), previous_sum, factors, dca_info_chunk[i]['sqmi']) \
            for x in filter_hard_area()))
    # remove no benefit cases 
    filter_benefit = lambda: it.compress(filter_hard_area(), benefit)
    # remove cases where high habitat is increased with more water
    for k in overloaded_hab:
        overload = \
                it.starmap(func.overload_increase, \
                ((np.array(x), previous_sum, k, factors, \
                dca_info_chunk[i]['MP Acres']) \
                for x in filter_benefit()))
    filter_overload = it.compress(filter_benefit(), overload)
    allowed_scenario.append(filter_benefit())

test_assign = it.product(*allowed_scenario[0:2])
array_assign = it.imap(np.array, test_assign)
stack_assign = it.imap(np.vstack, array_assign)
previous_stack = np.vstack(previous_case_chunk[0:2])
dca_info_stack = dca_info_chunk[0].append(dca_info_chunk[1])

hard_check = \
    it.starmap(func.hard_area_check, \
    ((x, previous_stack, soft_indices, dca_info_stack['sqmi'], 3) \
    for x in stack_assign))
# only keep those scenarios < 3 sq miles hard transition
filter_hard_area = it.compress(stack_assign, hard_check)



t0 = time()
a = func.countup(allowed_scenario[3])
t1 = time()
print t1 - t0
print a

# only consider DCAs in close proximity for change during single step
centers = dca_info.loc[inplay_matrix.index.tolist()].copy()
centers = centers.loc[:, ['Cent_N', 'Cent_E']]
dm = squareform(pdist(np.array(centers)))
filter_distance = 5000






vals = it.starmap(func.single_factor_total, \
        ((np.array(x), factors['bw'], areas) for x in filt))
filt_val =it.ifilterfalse(lambda x: x < 100, vals)

value_constraints = \
  [lambda x : compare_value(x, test_generic.water, test_areas, test_case, 1), \
  lambda x : compare_value(x, test_generic.bw, test_areas, test_case, 0.9),
  lambda x : compare_value(x, test_generic.mw, test_areas, test_case, 0.9),
  lambda x : compare_value(x, test_generic.pl, test_areas, test_case, 0.9),
  lambda x : compare_value(x, test_generic.ms, test_areas, test_case, 0.9),
  lambda x : compare_value(x, test_generic.md, test_areas, test_case, 0.9)]


pool = Pool(processes=4)
total_cases = pool.map(countup, filt)

def splitter(x, x_length, n_splits):
    split_size = -(-x_length/4)
    itter_list = []
    for i in range(0, n_splits):
        itter_list.append(it.islice(x, i*split_size, (i+1)*split_size))
    return itter_list

def apply_count(a):
    a += 1
    return a

pool = Pool(processes=4)
num = pool.map(func.countup, split_filter)







