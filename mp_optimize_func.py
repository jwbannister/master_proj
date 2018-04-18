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
    suffixes = ["_ac"] * 5 + ["_af/y"]
    value_columns = factors.columns + suffixes
    case_factors = pd.DataFrame(np.empty([len(case), len(factors.columns)]), \
            index=areas.index, columns=value_columns.tolist())
    for x in range(0, len(factors.columns)):
        case_factors.iloc[:, x] = case.dot(factors.iloc[:, x]) * areas
    return case_factors

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

def freeze_dcas(case, indices, check_case):
    case_array = np.array(case)
    comp = np.diag(case_array.dot(check_case.transpose())).tolist()
    return not all([comp[x]==1 for x in indices])

