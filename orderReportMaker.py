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

# read in full export
export = pd.read_csv(f'{download_folder}orders_export.csv', dtype={'Phone':str })

# convert 'created at' to timestamp called 'date ordered'
timestamps = pd.to_datetime(export['Created at'])
export['Date Ordered'] = timestamps.apply(lambda x: datetime(x.year, x.month, x.day))

#define today
now = datetime.now()
today = datetime(now.year, now.month, now.day)
today_string = today.strftime("%Y-%m-%d")

#Filter based on day of the week
  #if it's Monday, then get the past 2 days of orders
if today.weekday() ==0:
  sunday = export[export['Date Ordered'] == today - timedelta(days=1)]
  saturday = export[export['Date Ordered'] == today - timedelta(days=2)]
  orders = pd.concat([sunday, saturday], ignore_index = True)
else:
  orders = export[export['Date Ordered'] == today - timedelta(days=1)]

# Fill in missing values from row above to make up for prob with shopify export
orders = orders.fillna(method='ffill')

# Filter out pre-orders
orders = orders[~orders['Lineitem name'].str.contains('pre-order')]

# Filter out fulfilled
orders = orders[orders['Fulfillment Status'].str.contains('unfulfilled')]

# Shorten description for instore-pickup
orders.loc[orders['Shipping Method'].str.contains('in-store'),['Shipping Method']]= 'in-store pickup'

# Create dataframe for cutters
to_make = orders.copy()
to_make.loc[to_make['Shipping Method'].str.contains('Delivery|in-store'), 'Shipping Method'] = 'Delivery/in-store'
to_make = to_make[['Lineitem name','Name', 'Shipping Method']]
to_make = pd.pivot_table(to_make, index=['Lineitem name', 'Shipping Method'], aggfunc='count')

# Create dataframe for instore pickups
pick_ups = orders[orders['Shipping Method'].str.contains('in-store')][['Date Ordered',
                                                                      'Billing Name',
                                                                      'Lineitem name',
                                                                      'Shipping Method',
                                                                      'Email',
                                                                      'Phone' ]]

# Create dataframe for deliveries
deliveries = orders[orders['Shipping Method'].str.contains('Delivery')][['Date Ordered',
                                                            'Lineitem name',
                                                            'Shipping Method',
                                                            'Billing Name', 
                                                            'Shipping Street',
                                                            'Shipping City', 
                                                            'Shipping Zip',
                                                            'Shipping Phone', 
                                                            'Notes',]].sort_values(by=['Shipping Zip'])

# Create dataframe for shipping
to_ship = orders[orders['Shipping Method'].str.contains('UPS')][['Date Ordered',
                                                            'Lineitem name',
                                                            'Shipping Method',
                                                            'Billing Name',
                                                            'Shipping Name',
                                                            'Shipping Street',
                                                            'Shipping City', 
                                                            'Shipping Zip',
                                                            'Shipping Phone', 
                                                            'Notes',]]

# Print all dataframes to excel
with pd.ExcelWriter(f'{download_folder}Online Order Reports-{today_string}.xlsx') as writer:
    to_make.to_excel(writer, sheet_name='To Make', index=True)
    pick_ups.to_excel(writer, sheet_name='Pick Ups', index=False)
    deliveries.to_excel(writer,sheet_name='Deliveries', index=False)
    to_ship.to_excel(writer, sheet_name='To Ship', index=False)