import pandas as pd
import datetime
import math
from scipy.optimize import minimize
import decimal
import numpy as np
from datetime import date
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


def months_between(d1, d2):
    d1 = datetime.datetime.strptime(d1, "%Y-%m-%d")
    d2 = datetime.datetime.strptime(d2, "%Y-%m-%d")
    delta = relativedelta(d1, d2)
    return abs(delta.months + (delta.years * 12))


def get_next_coupon_date(d, period):
    next_coupon = datetime.datetime.strptime(d, "%Y-%m-%d") + relativedelta(months=period)
    return next_coupon


def calculate_sum_product_1(spread, row, baseline, remaining_period):
    np_around = 9

    # if remaining_period.is_integer():
    #     total_period = remaining_period + 1
    # else:
    #     total_period = math.floor(remaining_period) + 1

    # Expire date (Y) - Next coupon date (Y)
    # total_period = 0
    # total_period = row[7].year-row[11].year

    face_value = row.iloc[6]
    expire_date = row.iloc[7]
    coupon = row.iloc[9]
    frequency = row.iloc[10]
    next_coupon_date = row.iloc[11]

    cashflow_list = []
    discount_factor_list = []

    loop_to = 0
    coupon_months_plus = 0
    if (frequency == 1):
        # Expire date (Y) - Next coupon date (Y
        loop_to = (expire_date.year - (year-1))-1
        coupon_months_plus = 12

    if (frequency == 2):
        months_diff = months_between(expire_date.strftime('%Y-%m-%d'), str(year-1) + '-12-31')
        # loop_to = ((expire_date.year - (year-1))*2)-1
        loop_to = math.floor(months_diff/6)
        coupon_months_plus = 6

    for x in range(0, loop_to):
        coupon_date = get_next_coupon_date(
            str(next_coupon_date.year) + '-' + str(next_coupon_date.month) + '-' + str(next_coupon_date.day), (x*coupon_months_plus))

        t = excel_round(
            (days_between(coupon_date.strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)

        cashflow_norm = 0
        if (frequency == 1):
            cashflow_norm = coupon * face_value
        if (frequency == 2):
            cashflow_norm = (coupon * face_value) / 2

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
        (days_between(expire_date.strftime('%Y-%m-%d'), str(year-1) + '-12-31')) / 365.25, np_around)

    baseline_fin = 0
    if isinstance(remaining_period, int):
        baseline_fin = baseline[remaining_period]

    if isinstance(remaining_period, float):
        baseline_fin = baseline[math.ceil(remaining_period)]

    cashflow_fin = 0
    if (frequency == 1):
        cashflow_fin = (coupon * face_value) + face_value
    if (frequency == 2):
        cashflow_fin = (coupon * face_value)/2 + face_value

    discount_factor_fin_plus = np.around((1/(1 + baseline_fin + spread)), np_around)
    discount_factor_fin = np.around((discount_factor_fin_plus[0] ** rem_period), np_around)

    cashflow_list.append(cashflow_fin)
    discount_factor_list.append(discount_factor_fin)

    outSceen(str(expire_date) + ' - ' + str(rem_period) + ' - ' + str(cashflow_fin) + ' - ' + str(100 *
             discount_factor_fin) + ' - ' + str(baseline_fin) + ' - ' + str(discount_factor_fin_plus[0]))

    # Calculate the SUMPRODUCT
    sum_product = sum(a * b for a, b in zip(cashflow_list, discount_factor_list))

    return sum_product


def objective(spread, row, baseline, remaining_period, desired_sum_product, frequency):
    if frequency == 1 or frequency == 2:
        current_sum_product = calculate_sum_product_1(spread, row, baseline, remaining_period)

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
    # > 0 1 64 23 -  39 -- 48 49
    if index == 49:
        outSceen(row.iloc[1])
        outSceen(row.iloc[6])
        outSceen(row.iloc[8])
        outSceen(str(year-1) + '-12-31')
        outSceen(row.iloc[7])
        outSceen(row.iloc[9])
        outSceen(row.iloc[10])
        outSceen(row.iloc[12])
        outSceen('-----------------------')

        if row.iloc[10] == 1 or row.iloc[10] == 2:
            # Initial guess for spread
            initial_spread = 0

            # optimal_spread = 0
            # current_sum_product = calculate_sum_product_1(
            #     0, row, baseline, row.iloc[12])

            # Minimize the objective function
            remaining_period = row.iloc[13]
            market_value = row.iloc[8]
            frequency = row.iloc[10]
            result = minimize(objective, initial_spread, args=(
                row, baseline, remaining_period, market_value, frequency))
            optimal_spread = result.x[0]
            outSceen("Optimal spread:" + str(optimal_spread*100))
            calculated_spread.append(optimal_spread*100)
        else:
            calculated_spread.append('')


# df.insert(37, "Calculated Spreads", calculated_spread, True)
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
