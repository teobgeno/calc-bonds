import pandas as pd
import datetime
import math
from scipy.optimize import minimize
import decimal
import numpy as np
from datetime import date
from dateutil.relativedelta import relativedelta
from dateutil.relativedelta import relativedelta
# np.set_printoptions(precision=32)
# Face value             row[6]
# MV                     row[8]
# valuation date         end of year
# Expire date            row[7]
# Coupon                 row[9]
# Frequency              row[10]
# Next coupon date      (Expire date - valuation date)/365,25
# Next coupon date       row[11]
# Remaining period       row[12]
# decimal.getcontext().prec = 9
# np.set_printoptions(precision=8)


def outSceen(s):
    print(s)
    return 1


def excel_round(number, digits):
    # Similar to Excel, if the number is missing return 0
    if number is None or pd.isna(number):
        return 0

    context = decimal.getcontext()
    context.rounding = decimal.ROUND_HALF_UP
    number = decimal.Decimal(number)
    rounded_number = round(number, digits)
    return float(rounded_number)


def days_between(d1, d2):
    d1 = datetime.datetime.strptime(d1, "%Y-%m-%d")
    d2 = datetime.datetime.strptime(d2, "%Y-%m-%d")
    return abs((d1 - d2).days)


def get_next_coupon_date(d, period):
    next_coupon = datetime.datetime.strptime(d, "%Y-%m-%d") + relativedelta(months=period)
    return next_coupon


def calculate_sum_product_0(spread, row, baseline, remaining_period):
    # remaining_period = 0.09035
    rem_period = (days_between(row.iloc[7].strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25
    cashflow_list = []
    discount_factor_list = []
    baseline_fin = baseline[1]
    cashflow_fin = row.iloc[6]
    discount_factor_fin_plus = np.around((1/(1 + baseline_fin + spread)), 9)
    discount_factor_fin = np.around((discount_factor_fin_plus[0] ** rem_period), 9)
    # discount_factor_fin = 1/pow((1 + baseline_fin + spread),remaining_period)

    cashflow_list.append(cashflow_fin)
    discount_factor_list.append(discount_factor_fin)

    # Calculate the SUMPRODUCT
    sum_product = sum(a * b for a, b in zip(cashflow_list, discount_factor_list))

    outSceen('##')
    outSceen(spread)
    outSceen(discount_factor_fin)
    outSceen(str(remaining_period) + ' - ' + str(cashflow_fin) + ' - ' +
             str(100*discount_factor_fin) + ' - ' + str(baseline_fin))
    outSceen(str(sum_product))
    r = np.around((1/(1 + baseline_fin + spread)), 9)

    r2 = (r[0] ** remaining_period)*100
    r3 = excel_round(r2, 4)
    r4 = (0.967520342 ** 0.090349076)

    outSceen('d')
    outSceen(rem_period)
    outSceen('d')
    outSceen('+')
    outSceen(str(discount_factor_fin_plus[0]))
    outSceen(str(r[0]))
    outSceen('+')
    outSceen('^')
    outSceen(str(remaining_period))
    outSceen(str(discount_factor_fin))
    outSceen(str(r4))
    # outSceen(str(r[0]))
    # outSceen(str(r2))

    outSceen('^')
    # outSceen(str(r2))
    # outSceen(str(100*discount_factor_fin))
    outSceen('@@')

    return sum_product


def calculate_sum_product_1_cl(spread, row, baseline, remaining_period):
    np_around = 6
    rem_period = excel_round(
        (days_between(row.iloc[7].strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)

    # if remaining_period.is_integer():
    #     total_period = remaining_period + 1
    # else:
    #     total_period = math.floor(remaining_period) + 1

    # Expire date (Y) - Next coupon date (Y)
    # total_period = 0
    # total_period = row[7].year-row[11].year

    cashflow_list = []
    discount_factor_list = []

    loop_count = 0
    for x in range(row.iloc[11].year, row.iloc[7].year):

        # print(date(row.iloc[11].year+loop_count, row.iloc[11].month, row.iloc[11].day))
        coupon_date = date(row.iloc[11].year+loop_count, row.iloc[11].month, row.iloc[11].day)
        t = excel_round(
            (days_between(coupon_date.strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)

        cashflow_norm = row.iloc[9] * row.iloc[6]  # Coupon * Face value

        if isinstance(t, int):
            baseline_sel = baseline[t]

        if isinstance(t, float):
            baseline_sel = baseline[math.ceil(t)]

        discount_factor_norm_plus = excel_round((1/(1 + (baseline_sel) + spread)), np_around)
        discount_factor_norm = excel_round((discount_factor_norm_plus ** t), np_around)

        cashflow_list.append(cashflow_norm)
        discount_factor_list.append(discount_factor_norm)
        outSceen(str(x) + ' - ' + str(t) + ' - ' + str(cashflow_norm) + ' - ' + str(100*discount_factor_norm) +
                 '%' + ' - ' + str(baseline_sel) + ' - ' + str(discount_factor_norm_plus))
        loop_count = loop_count + 1

    # final line
    baseline_fin = 0
    if isinstance(remaining_period, int):
        baseline_fin = baseline[remaining_period]

    if isinstance(remaining_period, float):
        baseline_fin = baseline[math.ceil(remaining_period)]

    cashflow_fin = (row.iloc[9] * row.iloc[6]) + row.iloc[6]  # (Coupon *  Face value) + Face value
    # discount_factor_fin = 1/pow((1 + baseline_fin + spread),remaining_period)
    discount_factor_fin_plus = excel_round((1/(1 + baseline_fin + spread)), np_around)
    discount_factor_fin = excel_round((discount_factor_fin_plus ** rem_period), np_around)

    cashflow_list.append(cashflow_fin)
    discount_factor_list.append(discount_factor_fin)
    outSceen(str(remaining_period) + ' - ' + str(cashflow_fin) + ' - ' + str(100 *
             discount_factor_fin) + ' - ' + str(baseline_fin) + ' - ' + str(discount_factor_fin_plus))

    # Calculate the SUMPRODUCT
    sum_product = sum(a * b for a, b in zip(cashflow_list, discount_factor_list))

    outSceen('d')
    outSceen(rem_period)
    outSceen('d')

    return sum_product


def calculate_sum_product_1(spread, row, baseline, remaining_period):
    np_around = 9

    # if remaining_period.is_integer():
    #     total_period = remaining_period + 1
    # else:
    #     total_period = math.floor(remaining_period) + 1

    # Expire date (Y) - Next coupon date (Y)
    # total_period = 0
    # total_period = row[7].year-row[11].year

    cashflow_list = []
    discount_factor_list = []

    loop_to = 0
    coupon_months_plus = 0
    if (row.iloc[10] == 1):
        loop_to = row.iloc[7].year - row.iloc[11].year
        coupon_months_plus = 12
    if (row.iloc[10] == 2):
        loop_to = (row.iloc[7].year - row.iloc[11].year)*2
        coupon_months_plus = 6

    for x in range(0, loop_to):
        print(str(x))
        coupon_date = get_next_coupon_date(
            str(row.iloc[11].year) + '-' + str(row.iloc[11].month) + '-' + str(row.iloc[11].day), (x*coupon_months_plus))
        # coupon_date = date(row.iloc[11].year+loop_count, row.iloc[11].month, row.iloc[11].day)
        t = excel_round(
            (days_between(coupon_date.strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)

        cashflow_norm = row.iloc[9] * row.iloc[6]  # Coupon * Face value

        if isinstance(t, int):
            baseline_sel = baseline[t]

        if isinstance(t, float):
            baseline_sel = baseline[math.ceil(t)]

        discount_factor_norm_plus = np.around((1/(1 + (baseline_sel) + spread)), np_around)
        discount_factor_norm = np.around((discount_factor_norm_plus[0] ** t), np_around)

        cashflow_list.append(cashflow_norm)
        discount_factor_list.append(discount_factor_norm)
        outSceen(str(coupon_date) + ' - ' + str(t) + ' - ' + str(cashflow_norm) + ' - ' + str(100*discount_factor_norm) +
                 '%' + ' - ' + str(baseline_sel) + ' - ' + str(discount_factor_norm_plus[0]))

    # final line
    # (Expire date - Valuadtion date) / 365.25
    rem_period = excel_round(
        (days_between(row.iloc[7].strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)
    baseline_fin = 0
    if isinstance(remaining_period, int):
        baseline_fin = baseline[remaining_period]

    if isinstance(remaining_period, float):
        baseline_fin = baseline[math.ceil(remaining_period)]

    cashflow_fin = (row.iloc[9] * row.iloc[6]) + row.iloc[6]  # (Coupon *  Face value) + Face value
    discount_factor_fin_plus = np.around((1/(1 + baseline_fin + spread)), np_around)
    discount_factor_fin = np.around((discount_factor_fin_plus[0] ** rem_period), np_around)

    cashflow_list.append(cashflow_fin)
    discount_factor_list.append(discount_factor_fin)

    outSceen(str(row.iloc[7]) + ' - ' + str(rem_period) + ' - ' + str(cashflow_fin) + ' - ' + str(100 *
             discount_factor_fin) + ' - ' + str(baseline_fin) + ' - ' + str(discount_factor_fin_plus[0]))

    # Calculate the SUMPRODUCT
    sum_product = sum(a * b for a, b in zip(cashflow_list, discount_factor_list))

    return sum_product


def objective(spread, row, baseline, remaining_period, desired_sum_product, frequency):
    if frequency == 1:
        current_sum_product = calculate_sum_product_1(
            spread, row, baseline, remaining_period)

    if frequency == 0:
        current_sum_product = calculate_sum_product_0(spread, row, baseline, remaining_period)

    outSceen(str(current_sum_product) + ' ---' + str(desired_sum_product))
    return (current_sum_product - desired_sum_product) ** 2


'''
def objective_brute(spread, row, baseline, remaining_period, desired_sum_product):
    outSceen(spread)
    current_sum_product = calculate_sum_product(spread, row, baseline, remaining_period)
    if current_sum_product > row.iloc[8]:
        spread += 0.001
        objective_brute(spread, row, baseline, remaining_period, desired_sum_product)
    elif current_sum_product < row.iloc[8]:
        spread -= 0.001
        objective_brute(spread, row, baseline, remaining_period, desired_sum_product)
    else:
        outSceen('ok calculate spread')
        outSceen(spread)
'''

########### START ###########

import_file = "materials/import_bond_value.xlsx"

# baseline data
baseline = []
ds = pd.read_excel(import_file, sheet_name="RFR", usecols='B')
for index, row in ds.iterrows():
    if (isinstance(row.iloc[0], int) or isinstance(row.iloc[0], float)):
        baseline.append(row.iloc[0])

# read by default 1st sheet of an excel file
df = pd.read_excel(import_file, sheet_name="Bonds")

today = datetime.date.today()
year = today.year
calculated_spread = []
calculated_spread.append('')

for index, row in df.iterrows():
    # 1 64
    if index == 1:
        # outSceen(row.iloc[6])
        # outSceen(row.iloc[8])
        # outSceen(str(year-1) + '-12-31')
        # outSceen(row.iloc[7])
        # outSceen(row.iloc[9])
        # outSceen(row.iloc[10])
        # rem_period = (days_between(row.iloc[7].strftime(
        #     '%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25
        # outSceen('-----------------------')

        if row.iloc[10] == 1 or row.iloc[10] == 0:
            # Initial guess for spread
            initial_spread = 0

            # current_sum_product = calculate_sum_product_1_cl(0, row, baseline, row.iloc[12])

            # Minimize the objective function
            result = minimize(objective, initial_spread, args=(
                row, baseline, row.iloc[12], row.iloc[8], row.iloc[10]))
            optimal_spread = result.x[0]
            outSceen("Optimal spread:" + str(optimal_spread*100))
            calculated_spread.append(optimal_spread*100)
        else:
            calculated_spread.append('')


# df.insert(12, "Calculated Spreads", calculated_spread, True)
# file_name = 'calculated_spreads.xlsx'
# df.to_excel(file_name)
# print('done!!!')


# print(dataframe1)


# Define the values
# value_1 = 12.375
# percentage_value = 91404

# Convert the percentage to a decimal by dividing by 100
# decimal_percentage = percentage_value / 100

# Calculate the sum product
# sum_product = value_1 + decimal_percentage

# print(sum_product)
