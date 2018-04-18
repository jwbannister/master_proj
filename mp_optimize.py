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

# BUILD TEST DATA
info_file = pd.ExcelFile("/home/john/code/master_proj/DCA-DCM Constraints.xlsx")
manual = info_file.parse(sheet_name="Manual Scenario", header=2)
num_dca = 6
num_dcm = 7
base_case = manual.iloc[0:num_dca, 0:num_dcm]
factors = generic_factors.iloc[0:num_dcm, :].copy()
areas = assignments.area_ac[0:num_dca].copy()
previous_step = base_case.copy()

# read DCA-DCM constraints file
constraints = info_file.parse(sheet_name="Constraints", skiprows=8)
# reduce contraints table to match test data
constraints = constraints.iloc[0:num_dca, 0:num_dcm]
dca_info = pd.read_csv("~/code/master_proj/DCA_top_detailed_MP.csv", header=0)
dca_info.drop_duplicates(subset='MP Name', inplace=True)
dca_info.index = [x.strip() for x in dca_info['MP Name']]
phase_info = base.join(dca_info.loc[:, 'Phase'])
frozen_dcas = phase_info.loc[phase_info['Phase'].\
        isin(['7A', '7B', '8', '9', '10'])].index.tolist()
frozen_dcas.extend([u'Channel Area North', u'Channel Area South'])

# split off matrix of DCAs that cannot change DCM ("frozen")
frozen_dcas = [x for x in frozen_dcas if x in areas.index.tolist()]
frozen_indices = [areas.index.tolist().index(x) for x in frozen_dcas]
frozen_matrix = previous_step.iloc[frozen_indices]
# create matrix of DCAs that can change in this step ("inplay")
inplay_dcas = [x for x in areas.index.tolist() if x not in frozen_dcas]
inplay_indices = [areas.index.tolist().index(x) for x in inplay_dcas]
inplay_matrix = previous_step.iloc[inplay_indices]

# set contraints on which DCAs can by which DCMs
dcm_constraints = pd.read_csv("~/code/master_proj/dcm_constraints.csv", \
        header=0, index_col='dca')
dcm_iter = []
for i in range(0, len(dcm_constraints)):
    tmp = dcm_constraints.iloc[i].tolist()
    tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
    a = []
    for d in tmp_ind:
        b = [0 for x in tmp]
        b[d] = 1
        a.append(b)
    tuple(a)
    dcm_iter.append(a)








# only consider DCAs in close proximity for change during single step
centers = dca_info.loc[inplay_matrix.index.tolist()].copy()
centers = centers.loc[:, ['Cent_N', 'Cent_E']]
dm = squareform(pdist(np.array(centers)))
filter_distance = 5000



# matrix of all possible DCM assignment vectors
identity_assignments = np.identity(len(factors)).tolist()
identity_assignments = [tuple(identity_assignments)] * len(inplay_matrix)
assignment_combinations = lambda: it.product(*identity_assignments)
combos = [np.array(x) for x in assignment_combinations()]

filter1 = it.ifilterfalse(precalc_constraints[0], assignment_combinations())

t0 = time()
n = func.countup(assignment_combinations())
t1 = time()
print t1 - t0

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










