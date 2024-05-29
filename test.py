import numpy as np
import pandas as pd
from scipy.optimize import minimize

# Example data
row = pd.Series([0, 0, 0, 0, 0, 0, 0, 1.2, 0, 100])
baseline = [0, 0.02, 0.03, 0.04, 0.05]
total_period = 5
desired_sum_product = 1523  # The desired value for sum_product

def calculate_sum_product(spread, row, baseline, total_period):
    cashflow_list = []
    discount_factor_list = []
    
    for x in range(1, total_period):
        cashflow_norm = row.iloc[9] * row.iloc[6]
        discount_factor_norm = 1 / pow((1 + baseline[x] + spread), x)
        cashflow_list.append(cashflow_norm)
        discount_factor_list.append(discount_factor_norm)
    
    sum_product = sum(a * b for a, b in zip(cashflow_list, discount_factor_list))
    print(sum_product)
    return sum_product

# Objective function to minimize the difference between current and desired sum_product
def objective(spread, row, baseline, total_period, desired_sum_product):
    current_sum_product = calculate_sum_product(spread, row, baseline, total_period)
    return (current_sum_product - desired_sum_product) ** 2

# Initial guess for spread
initial_spread = 0

# Minimize the objective function
result = minimize(objective, initial_spread, args=(row, baseline, total_period, desired_sum_product), method='BFGS')

optimal_spread = result.x[0]
print("Optimal spread:", optimal_spread)