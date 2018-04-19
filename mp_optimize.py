import pandas as pd
import numpy as np
import itertools as it
from tqdm import tqdm
import csv
from multiprocessing import Pool
import mp_optimize_func as func
from time import time
from scipy.spatial.distance import pdist, squareform

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
constraints = info_file.parse(sheet_name="Constraints", header=8)

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
previous_case = np.array(step0.iloc[0:num_dca, 0:num_dcm]).copy()
# reduce contraints table to match test data
constraints = np.array(constraints.iloc[0:num_dca, 0:num_dcm]).copy()
# reduce DCA info to match test data
dca_info = dca_info.iloc[0:num_dca, 0:num_dcm].copy()
dca_info = dca_info.loc[:, ['MP Acres', 'Cent_N', 'Cent_E']]
dca_info['sqmi'] = dca_info['MP Acres'] * 0.0015625

# evaluate base case habitat and water usage
base_total = func.evaluate_case(base_case, factors, dca_info['MP Acres']).sum()
# where is current scenario compared to base?
previous_total = func.evaluate_case(previous_case, factors, dca_info['MP Acres']).sum()
previous_percent = previous_total/base_total
overloaded_hab = [y for x, y in enumerate(previous_percent.index.tolist()) \
        if previous_percent[x] > 1]
underloaded_hab = [y for x, y in enumerate(previous_percent.index.tolist()) \
        if previous_percent[x] < 0.9]

# which DCMs count as a "soft" transition?
soft_dcms = ['Tillage', 'Brine', 'Till-Brine']
soft_indices = [x for x, y in enumerate(factors.index.tolist()) if y in soft_dcms]

# split data into chunks for processing
chunk_size = 160
num_chunks = -(-num_dca // chunk_size)
base_case_chunk = [base_case[chunk_size*n:chunk_size*(n+1)] \
        for n in range(0, num_chunks)]
previous_case_chunk = [previous_case[chunk_size*n:chunk_size*(n+1)] \
        for n in range(0, num_chunks)]
constraints_chunk = [constraints[chunk_size*n:chunk_size*(n+1)] \
        for n in range(0, num_chunks)]
dca_info_chunk = [dca_info.iloc[chunk_size*n:chunk_size*(n+1)] \
        for n in range(0, num_chunks)]

allowed_scenario = []
for i in range(0, num_chunks):
    # build list of allowable assignments for each DCA
    allowed_cases = []
    for j in range(0, len(constraints_chunk[i])):
        tmp = constraints_chunk[i][j].tolist()
        tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
        a = []
        for d in tmp_ind:
            b = [0 for x in tmp]
            b[d] = 1
            a.append(b)
        allowed_cases.append(a)
    n_allowed = reduce(lambda x, y: x*y, [len(z) for z in allowed_cases])

    smart_cases = []
    for dca in range(0, len(allowed_cases)):
        dca_assigns = []
        for case in range(0, len(allowed_cases[dca])):
            if func.evaluate_dca_change(allowed_cases[dca][case], \
                    previous_case_chunk[i][dca], factors, overloaded_hab, \
                    underloaded_hab):
                dca_assigns.append(allowed_cases[dca][case])
        smart_cases.append(dca_assigns)



    


    previous_sum = func.evaluate_case(previous_case_chunk[i], factors, \
            dca_info_chunk[i]['MP Acres']).sum()



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







