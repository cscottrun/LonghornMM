# import dependancies
import os
download_folder = os.path.expanduser("~")+"/Downloads/"
import calendar
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter
import openpyxl

# read in full export
export = pd.read_csv(f'{download_folder}orders_export.csv', dtype={'Phone':str })

# convert 'created at' to timestamp called 'date ordered'
timestamps = pd.to_datetime(export['Created at'])
export['Date Ordered'] = timestamps.apply(lambda x: datetime(x.year, x.month, x.day))

#define today
now = datetime.now()
today = datetime(now.year, now.month, now.day)
today_string = today.strftime("%Y-%m-%d")

orders = export
# filter out orders from today
orders = orders[orders['Date Ordered'] != today]
# orders = orders[orders['Date Ordered'] == today]

# Fill in missing values for SOME columns with info from row above
orders['Fulfillment Status'] = orders['Fulfillment Status'].fillna(method='ffill')
orders['Shipping Name'] = orders['Shipping Name'].fillna(method='ffill')
orders['Billing Name'] = orders['Billing Name'].fillna(method='ffill')
orders['Shipping Street'] = orders['Shipping Street'].fillna(method='ffill')
orders['Shipping City'] = orders['Shipping City'].fillna(method='ffill')
orders['Shipping Province'] = orders['Shipping Province'].fillna(method='ffill')
orders['Shipping Zip'] = orders['Shipping Zip'].fillna(method='ffill')
orders['Shipping Country'] = orders['Shipping Country'].fillna(method='ffill')
orders['Shipping Method'] = orders['Shipping Method'].fillna(method='ffill')
orders['Financial Status'] = orders['Financial Status'].fillna(method='ffill')

# Filter out pre-orders
#orders = orders[~orders['Lineitem name'].str.contains('pre-order')]

# Filter out fulfilled
orders = orders[~orders['Fulfillment Status'].str.match('fulfilled')]

# Filter to only get 'paid' (to get rid of refunded)
orders = orders[~orders['Financial Status'].str.match('refunded')]

# Shorten description for instore-pickup
orders.loc[orders['Shipping Method'].str.contains('in-store'),['Shipping Method']]= 'in-store pickup'

# Shipping Method Summary
summary = orders.copy()
summary.loc[summary['Shipping Method'].str.contains('Delivery|Austin|shipping'), 'Shipping Method'] = 'Delivery'
summary.loc[summary['Shipping Method'].str.contains('pickup'), 'Shipping Method'] = "Pick Up"
summary = summary[['Lineitem name','Name', 'Shipping Method', 'Lineitem quantity']]
summary = pd.pivot_table(summary, index=['Shipping Method', 'Lineitem name'], values='Lineitem quantity', aggfunc=np.sum)

# Create list for cutters
to_make = orders.copy()
to_make = pd.pivot_table(to_make, index=['Lineitem name'], values='Lineitem quantity', aggfunc=np.sum)

# Create dataframe for instore pickups
pick_ups = orders[orders['Shipping Method'].str.contains('in-store')][['Name',
                                                                       'Date Ordered',
                                                                       'Shipping Name',
                                                                       'Lineitem name',
                                                                       'Lineitem quantity',
                                                                       'Shipping Method',
                                                                       'Email',
                                                                       'Phone']]

# Create dataframe for deliveries
deliveries = orders[orders['Shipping Method'].str.contains('Delivery')][['Name',
                                                            'Date Ordered',
                                                            'Lineitem name',
                                                            'Lineitem quantity',
                                                            'Shipping Method',
                                                            'Shipping Name', 
                                                            'Shipping Street',
                                                            'Shipping City', 
                                                            'Shipping Zip',
                                                            'Shipping Phone', 
                                                            'Notes',]].sort_values(by=['Shipping Method', 'Name'])

# Create dataframe for shipping
to_ship = orders[orders['Shipping Method'].str.contains('UPS')][['Name',
                                                            'Date Ordered',
                                                            'Lineitem name',
                                                            'Lineitem quantity',
                                                            'Shipping Method',
                                                            'Shipping Name',
                                                            'Shipping Name',
                                                            'Shipping Street',
                                                            'Shipping City', 
                                                            'Shipping Zip',
                                                            'Shipping Phone', 
                                                            'Notes',]]

# Print all dataframes to excel
with pd.ExcelWriter(f'{download_folder}Online Order Reports-{today_string}.xlsx') as writer:
    summary.to_excel(writer, sheet_name='Method Summary')
    to_make.to_excel(writer, sheet_name='To Make')
    pick_ups.to_excel(writer, sheet_name='Pick Ups', index=False)
    deliveries.to_excel(writer,sheet_name='Deliveries', index=False)
    to_ship.to_excel(writer, sheet_name='To Ship', index=False)

# print unique delivery number
delNum=deliveries['Shipping Name'].nunique()
print(f'Unique Deliveries: {delNum}')                                                                                                                       