"""

A class built to handle the storage of product objects which may store various attributes of the product's
sales, quantities etc.

"""

# from script import TimePeriod
# from script import time_period
# Track what type of data user wants to format
from enum import Enum


class TimePeriod(Enum):
    MONTH = "M"
    ANNUAL = "A"

class Product:
    def __init__(self, sku):
        self._sku = sku
        self._description = None
        self._total_quantity = 0
        self._total_sales = 0
        self._total_profit = 0
        self._total_transactions = [] # store invoice dates
        self._documents = []

        self._transaction_years = {}
        self._transaction_lists = [] # may be used later to track document numbers for transactions in each year etc
        self._quantity_lists = []
        self._sales_lists = []
        self._profit_lists = []

    # Getters and Setters for all private variables

    def get_sku(self) -> str:
        return self._sku
    
    def get_description(self) -> str:
        return self._description

    def set_description(self, desc: str):
        self._description = desc

    def get_total_quantity(self) -> int:
        return self._total_quantity
    
    def increment_total_quantity(self, amount: int):
        self._total_quantity += amount

    def get_total_sales(self) -> int:
        return self._total_sales
    
    def increment_total_sales(self, amount: int):
        self._total_sales += amount

    def get_total_profit(self) -> int:
        return self._total_profit
    
    def increment_total_profit(self, amount: int):
        self._total_profit += amount

    def get_all_transactions(self) -> list:
        return self._total_transactions
    
    def get_all_transaction_lists(self) -> list:
        return self._transaction_lists
    
    def get_all_quantity_lists(self) -> list:
        return self._quantity_lists
    
    def get_all_sales_lists(self) -> list:
        return self._sales_lists
    
    def get_all_profit_lists(self) -> list:
        return self._profit_lists

    def update_product(self, quantity, sales, profit, invoice_date, document, time_period):
        """Increment product totals with new row data"""
        self.increment_total_quantity(quantity)
        self.increment_total_sales(sales)
        self.increment_total_profit(profit)
        self._total_transactions.append(invoice_date) # append all dates to one list
        self._documents.append(document)


        # Handle organising transactions by year v. and then month?
        # Get year
        year = str(invoice_date)[-4:] # last 4 characters in date i.e. YYYY btw date is float DD-APR-YYYY
        # Get month
        month = str(invoice_date)[3:6] 
        month = month + ' ' + year 

        if time_period == "Month":
            value_to_sort_by = month
        else: # annual
            value_to_sort_by = year # change based on user input eventually

        if value_to_sort_by in self._transaction_years: # if year already discovered, increment totals for that year
            index = self._transaction_years[value_to_sort_by]
            self._sales_lists[index] += sales
            self._quantity_lists[index] += quantity
            self._profit_lists[index] += profit
        else: # year not discovered, begin totals for new year
            index = len(self._transaction_years) # get next index
            self._transaction_years[value_to_sort_by] = index # store year and index
            self._sales_lists.append(sales)
            self._quantity_lists.append(quantity)
            self._profit_lists.append(profit)

    def __str__(self):
        return (f"Product SKU: {self._sku}, "
                f"Total Quantity: {self._total_quantity}, "
                f"Total Sales: {self._total_sales}, "
                f"Total Profit: {self._total_profit}, ")