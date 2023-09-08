### IMPORTING PACKAGES

import pandas as pd
import numpy as np
from math import ceil


# Read excel from file locations
path1 = 'C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\Assignment details\\Company X - Order Report.xlsx'
path2 = 'C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\Assignment details\\Company X - Pincode Zones.xlsx'
path3 = 'C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\Assignment details\\Company X - SKU Master.xlsx'
path4 = 'C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\Assignment details\\Courier Company - Invoice.xlsx'
path5 = 'C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\Assignment details\\Courier Company - Rates.xlsx'

df_Orders = pd.read_excel(path1)
df_pincodes = pd.read_excel(path2)
df_sku = pd.read_excel(path3)
df_courierInvoice = pd.read_excel(path4)
df_courierRates = pd.read_excel(path5)


### CALULATE TOTAL WEIGHT per SKU, COD PRICE
df_Orders_sum = df_Orders.merge(df_sku, on='SKU', how='left')
df_Orders_sum['Net Weight(kg)'] = (df_Orders_sum['Order Qty']*df_Orders_sum['Weight (g)']/1000).round(2)
df_Orders_sum['Net Price'] = df_Orders_sum['Item Price(Per Qty.)']

for i in range(df_Orders_sum.shape[0]):
  df_Orders_sum['Payment Mode'][i] = (0, 1)[ df_Orders_sum['Payment Mode'][i] == 'COD' ] ### mapping PREPAID-COD to 0-1 for calculative convinience

for i in range(df_Orders_sum.shape[0]):
  if df_Orders_sum['Payment Mode'][i]:                                          ### if PAYMENT MODE = 1, i.e COD
    price = df_Orders_sum['Net Price'][i]
    if price > 300:
      df_Orders_sum['Net Price'][i] = 0.05*df_Orders_sum['Net Price'][i]
    else:
      df_Orders_sum['Net Price'][i] = 15
  else:                                                                         ### if PAYMENT MODE = 0, i.e PREPAID
    df_Orders_sum['Net Price'][i] = 0

df_Orders_sum.head()



#TOTAL_WEIGHT and COD_CHARGES for each ORDER_ID
### ORDER_ID is the KEY to both the mapping

total_weights = {}
COD_charges = {}

for i in range(df_Orders_sum.shape[0]):
  id = df_Orders_sum['ExternOrderNo'][i]
  total_weights[id] = total_weights.get(id, 0) + df_Orders_sum['Net Weight(kg)'][i]
  COD_charges[id] = COD_charges.get(id, 0) +df_Orders_sum['Net Price'][i]



##### Mapping PINCODES provided by COMPANY X to that of the COURIER COMPANY
pincode_map = {}
for i in range(df_pincodes.shape[0]):
  pincode_map[df_pincodes['Customer Pincode'][i]] = df_pincodes['Zone'][i]
print(len(pincode_map))

df_courierInvoice['Zone by X'] = df_courierInvoice['Customer Pincode'].map(pincode_map)
df_courierInvoice.head()



### creating a Collective Order details consisting data from all the tables
df_order_details = pd.DataFrame()
df_order_details[['Order ID', 'AWB Number', 'Zone (Courier Company)', 'Zone (as per X)','Weight(courier comp.)', 'Billing Amount(Rs.)']] = df_courierInvoice[['Order ID', 'AWB Code', 'Zone', 'Zone by X', 'Charged Weight', 'Billing Amount (Rs.)']]

df_order_details['Weight (as per X)'] = df_order_details['Order ID'].map(total_weights)
df_order_details['COD charge'] = df_order_details['Order ID'].map(COD_charges)

df_order_details.head()

weight_slabs = {'a':0.25, 'b':0.50, 'c':0.75, 'd':1.25, 'e':1.50}
df_courierRates.set_index('Zone', inplace=True)
df_courierRates.head()



### calculate weight slabs for courier company
df_courier_slabs = df_order_details[['Order ID', 'Zone (Courier Company)', 'Weight(courier comp.)']]
df_courier_slabs['num_slabs'] = [0]*len(df_courier_slabs)
df_courier_slabs['weight_slabs_byCourier'] = [0]*len(df_courier_slabs)

for i in range(df_courier_slabs.shape[0]):
  num = ceil(df_courier_slabs['Weight(courier comp.)'][i] / weight_slabs[df_courier_slabs['Zone (Courier Company)'][i]])
  df_courier_slabs['num_slabs'][i] = num
  df_courier_slabs['weight_slabs_byCourier'][i] = num * weight_slabs[df_courier_slabs['Zone (Courier Company)'][i]]


df_courier_slabs.head()

df_expected_charge_calc = df_order_details[['Order ID', 'Zone (as per X)', 'Weight (as per X)', 'COD charge']]
df_expected_charge_calc = df_expected_charge_calc.merge(df_courierInvoice[['Order ID', 'Type of Shipment']], on='Order ID', how='left')

df_expected_charge_calc['num_slabs'] = [0]*len(df_expected_charge_calc)
df_expected_charge_calc['weight_slabs'] = [0]*len(df_expected_charge_calc)
for i in range(df_expected_charge_calc.shape[0]):
  num = ceil(df_expected_charge_calc['Weight (as per X)'][i] / weight_slabs[df_expected_charge_calc['Zone (as per X)'][i]])
  df_expected_charge_calc['num_slabs'][i] = num
  df_expected_charge_calc['weight_slabs'][i] = num * weight_slabs[df_expected_charge_calc['Zone (as per X)'][i]]

df_expected_charge_calc.head()



###calculating total expected fare
forward_charge, rto_charge = [], []
for i in range(df_expected_charge_calc.shape[0]):
  id = df_expected_charge_calc['Zone (as per X)'][i].upper()
  extra_slabs = df_expected_charge_calc['num_slabs'][i]-1

  forward_charge.append(df_courierRates['Forward Fixed Charge'][id] + extra_slabs*df_courierRates['Forward Additional Weight Slab Charge'][id])
  if df_expected_charge_calc['Type of Shipment'][i] == 'Forward charges':
    rto_charge.append(0)

  else:
    rto_charge.append(df_courierRates['RTO Fixed Charge'][id] + extra_slabs*df_courierRates['RTO Additional Weight Slab Charge'][id])

df_expected_charge_calc['forward_charge'] = forward_charge
df_expected_charge_calc['rto_charge'] = rto_charge
df_expected_charge_calc['total_sum'] = df_expected_charge_calc['COD charge'] + df_expected_charge_calc['forward_charge'] + df_expected_charge_calc['rto_charge']
df_expected_charge_calc.head(10)

df_order_details = df_order_details.merge(df_expected_charge_calc[['Order ID', 'weight_slabs', 'total_sum']], on='Order ID', how='left')
df_order_details.head()

df_order_details = df_order_details.merge(df_courier_slabs[['Order ID', 'weight_slabs_byCourier']], on='Order ID', how='left')
df_order_details.head()

df_order_details['Difference'] = df_order_details['total_sum']-df_order_details['Billing Amount(Rs.)']
df_order_details.head()



### REORDERING COLUMNS
df_order_details = df_order_details.loc[:,['Order ID', 'AWB Number', 'Weight (as per X)', 'weight_slabs', 'Weight(courier comp.)', 'weight_slabs_byCourier',
                          'Zone (as per X)', 'Zone (Courier Company)', 'total_sum', 'Billing Amount(Rs.)', 'Difference']]

df_order_details.head()



### Preparing the SUMMARY DATAFRAME
count_corr, count_over, count_under = 0, 0, 0
amt_corr, amt_over, amt_under = 0, 0, 0
for i in range(df_order_details.shape[0]):
  if df_order_details['Difference'][i] == 0:
    count_corr += 1
    amt_corr += df_order_details['Billing Amount(Rs.)'][i]
  elif df_order_details['Difference'][i] > 0:
    count_under += 1
    amt_under += df_order_details['Difference'][i]
  elif df_order_details['Difference'][i] < 0:
    count_over += 1
    amt_over += df_order_details['Difference'][i]

df_summary = pd.DataFrame()
df_summary[''] = ['Total Orders - Correctly Charged', 'Total Orders - Over Charged', 'Total Order - Under Charged']
df_summary['Count'] = [count_corr, count_over, count_under]
df_summary['Amount'] = [amt_corr, amt_over, amt_under]

df_summary.head()



### RENAMING THE COLUMNS OF DATAFRAME
df_order_details.columns.values[:] = ['Order ID', 'AWB Number', 'Total weight as per X (KG)', 'Weight slab as per X (KG)', 'Total weight as per Courier Company (KG)', 'Weight slab charged by Courier Company (KG)', 'Delivery Zone as per X', 'Delivery Zone charged by Courier Company', 'Expected Charge as per X (Rs.)', 'Charges Billed by Courier Company (Rs.)', 'Difference Between Expected Charges and Billed Charges (Rs.)']
df_order_details.head()



####saving order details to EXCEL files on drive
df_order_details.to_excel('C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\cointab.xlsx', index=False)

df_summary.to_excel('C:\\Users\\Harshit Jha\\Downloads\\Cointab - Asessment\\SUBMISSION\\summary.xlsx', index=False)



