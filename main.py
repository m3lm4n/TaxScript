# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import json
from datetime import datetime, timedelta
from urllib import request
from math import isclose

from openpyxl import load_workbook


def check_nbp_price(date):
    prev_date = date - timedelta(days=1)
    day_string = prev_date.strftime("%Y-%m-%d")

    url = 'http://api.nbp.pl/api/exchangerates/rates/a/usd/%s/' % (day_string)

    try:
        with request.urlopen(url) as response:
            body = json.loads(response.read())
            return body['rates'][0]['mid']
    except:
        new_date = prev_date - timedelta(days=1)
        return check_nbp_price(new_date)


def parse_saxo_trades():
    transactions = []

    wb = load_workbook(filename='saxo_trades_executed.xlsx')

    trades = wb['Trades']

    itertrades = iter(trades.rows)
    next(itertrades)
    for row in itertrades:
        date = row[3].value
        instrument = row[2].value
        amount = row[6].value
        booked_amount = row[10].value

        transactions.append((instrument, date, amount, booked_amount, booked_amount / amount))

    return transactions


def parse_revolut_trades():
    transactions = []

    deposit = []
    withdrawal = []
    dividend = []
    fee = []
    split = []

    wb = load_workbook(filename='revolut_trades_executed.xlsx')

    trades = wb['in']

    itertrades = iter(trades.rows)
    next(itertrades)
    for row in itertrades:
        split_row = row[0].value.split(',')
        (date, instrument, operation, amount, per_share, total, currency, fx_rate) = split_row

        if operation == 'STOCK SPLIT':
            split.append(row)
            transactions.append(("SPLIT", instrument, date, float(amount)))
        elif operation == 'CASH TOP-UP':
            deposit.append((date, float(total)))
        elif operation == 'CUSTODY_FEE':
            fee.append((date, float(total)))
        elif operation == 'DIVIDEND':
            dividend.append((date, instrument, float(total)))
        elif operation == 'CASH WITHDRAWAL':
            withdrawal.append((date, float(total)))
        else:
            if operation == 'SELL':
                amount = -float(amount)
            transactions.append((instrument, datetime.strptime(date, "%d/%m/%Y %H:%M:%S"), float(amount), float(total),
                                 float(total) / float(amount)))

    dep_sum = 0
    for dep in deposit:
        dep_sum += dep[1]

    print("Revolut deposited total of: %.2f" % dep_sum)

    wit_sum = 0
    for wit in withdrawal:
        wit_sum += abs(wit[1])

    print("Revolut withdrew total of %.2f" % wit_sum)

    return transactions


def process_transactions(transactions):
    to_date = {}
    to_tax = []

    for transaction in transactions:
        if transaction[0] == "SPLIT":
            print("TESLA SPLIT")
            continue

        (instrument, date, amount, booked_amount, per_share) = transaction

        if instrument not in to_date:
            to_date[instrument] = (0, [])

        (amount_owned, price_stack) = to_date[instrument]

        if amount > 0:
            amount_owned += amount
            price_stack.append((amount, per_share, booked_amount))

            to_date[instrument] = (amount_owned, price_stack)
        else:
            print("Selling %d of %s" % (abs(amount), instrument))
            print(to_date[instrument])
            sold_amount = amount
            tax = 0
            while not isclose(amount, 0, abs_tol=0.000001):
                if len(price_stack) == 0:
                    print("------------------- Taxation event with 0 shares. Check for errors --------------------------")
                    break

                (trans_amount, trans_per_share, trans_booked_amount) = price_stack.pop(0)

                bought_at = trans_per_share
                sold_at = per_share
                per_share_taxable = abs(sold_at) - abs(bought_at)

                if abs(trans_amount) <= abs(amount):
                    amount += abs(trans_amount)

                    if amount > 0:
                        price_stack.insert(0, (amount, trans_per_share, amount * trans_per_share))

                    tax += per_share_taxable * abs(trans_amount)
                else:
                    new_trans_amount = abs(trans_amount) - abs(amount)
                    amount = 0
                    new_trans_booked_amount = new_trans_amount * trans_per_share

                    amount_sold = abs(trans_amount) - abs(new_trans_amount)
                    tax += per_share_taxable * amount_sold

                    price_stack.insert(0, (new_trans_amount, trans_per_share, new_trans_booked_amount))

            to_tax.append((sold_amount, instrument, per_share, tax))

            print("Taxation event sold %d of %s at %.2f per share. Taxable income is %.2f" % (
            abs(sold_amount), instrument, abs(per_share), tax))

    return to_tax


def apply_exchange_rate(transaction):
    if transaction[0] != "SPLIT":
        (instrument, date, amount, booked_amount, per_share) = transaction
        exchange_rate = check_nbp_price(date)

        return instrument, date, amount, booked_amount * exchange_rate, per_share * exchange_rate
    else:
        return transaction


if __name__ == '__main__':

    revolut_transaction_list = parse_revolut_trades()
    saxo_transaction_list = parse_saxo_trades()

    transaction_list = revolut_transaction_list + saxo_transaction_list
    transaction_list = map(apply_exchange_rate, transaction_list)

    taxable_events = process_transactions(transaction_list)

    final_taxable_income = 0
    for tax in taxable_events:
        final_taxable_income += tax[3]

    final_tax = final_taxable_income * 0.19
    print("Final taxable income %.2f" % final_taxable_income)
    print("Final tax %.2f" % final_tax)
