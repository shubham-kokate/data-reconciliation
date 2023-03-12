import pandas as pd
import numpy as np

#company-data
companyXOrderReport = pd.read_excel('input-data\Company X - Order Report.xlsx')
companyXPinCodes = pd.read_excel('input-data\Company X - Pincode Zones.xlsx')
companyXSKUMaster = pd.read_excel('input-data\Company X - SKU Master.xlsx')

#print(companyXSKUMaster.value_counts())
#'GIFTBOX202002' value is repeated. Since this is a Master List all values should only exisit once thus we have to remove duplicates
companyXSKUMaster =  companyXSKUMaster.drop_duplicates()

#courier-data
courierRates = pd.read_excel('input-data\Courier Company - Rates.xlsx')
courierInvoice = pd.read_excel('input-data\Courier Company - Invoice.xlsx')

#Declaring a WeightSlab Dictionary
weightslabs = {
    'a': 0.25,
    'b': 0.5,
    'c': 0.75,
    'd': 1.25,
    'e': 1.5
}

#Merging DataFrames on SKU Column to get a SQL type Join
SKUMerge = companyXOrderReport.merge(companyXSKUMaster, on='SKU', sort=True)
SKUMerge.rename(columns={
    'Weight (g)':'Weight per Item'
}, inplace=True)

#Calculating Total Weight Column based on Oty and Weight per Item
SKUMerge['Total Weight (g)'] = SKUMerge.apply(lambda x: x['Order Qty'] * x['Weight per Item'], axis=1)

#Getting Total Order Weight (g) column added to the current DataFrame
totalOrderWeight =  SKUMerge.groupby(by='ExternOrderNo')['Total Weight (g)'].sum()
totalOrderWeightDF = totalOrderWeight.to_frame()

totalOrderWeightDF['ExternOrderNo'] = totalOrderWeightDF.index
index = pd.Index(range(len(totalOrderWeightDF)))
totalOrderWeightDF.set_index(index, inplace=True)

totalOrderWeightDF.rename(columns={
    'Total Weight (g)': 'Total Order Weight (g)'
}, inplace=True)

SKUMerge = SKUMerge.merge(totalOrderWeightDF, on='ExternOrderNo')


colOrderID = SKUMerge['ExternOrderNo']
colTotalOrderWeight = SKUMerge['Total Order Weight (g)']

i = range(len(colOrderID))
OrderAgg = pd.DataFrame({
    'Order ID': colOrderID,
    'Total Order Weight (g)': colTotalOrderWeight
}, index=i)

OrderAgg = OrderAgg.drop_duplicates()

OrderAgg['Total Order Weight (kg)'] = OrderAgg.apply(lambda x: x['Total Order Weight (g)'] * 0.001, axis=1)

OrderAgg = OrderAgg.drop('Total Order Weight (g)', axis=1)

#Extracting Delivery Information from Courier Invoices to COMPARE PINCODES and ZONES
"""
courierCompanyWareHousePincode = courierInvoice['Warehouse Pincode']
courierCompanyCustomerPincode = courierInvoice['Customer Pincode']
courierCompanyZone = courierInvoice['Zone']

DeliveryInfo = pd.DataFrame({
    'Warehouse Pincode':courierCompanyWareHousePincode,
    'Customer Pincode': courierCompanyCustomerPincode,
    'Zone' : courierCompanyZone
})

#comparison = DeliveryInfo.eq(companyXPinCodes)
#print(comparison[comparison['Zone'] == True].count())

Out of 124 Values 65 Zones are marked Differently by Both Parties and 59 are marked same.
"""

MergedTable = courierInvoice.merge(OrderAgg, on='Order ID')

companyXPinCodes.rename(columns={
    'Zone': 'ZonebyX'
}, inplace=True)

MergedTable = MergedTable.merge(companyXPinCodes, how='left').drop_duplicates(keep='first', ignore_index=True)

#Checking if the Zones have filled according to previous comparison
"""
ZonebyCour = MergedTable['Zone']
ZonebyX = MergedTable['ZonebyX']
print(ZonebyCour.eq(ZonebyX).value_counts())
"""

#Calculating Applicable Weight According to X
def calcWeight(totalWeight, zone):
    applicableWeight = 0
    weightSlab = weightslabs[zone]
    while(totalWeight>0):
        applicableWeight = applicableWeight + weightSlab
        totalWeight = totalWeight - weightSlab
    return applicableWeight


for i in range(len(MergedTable)):
    MergedTable.loc[i, 'WeightbyX'] = calcWeight(MergedTable.loc[i, 'Total Order Weight (kg)'], MergedTable.loc[i, 'ZonebyX'])

for i in range(len(MergedTable)):
    MergedTable.loc[i, 'Weight slab charged by Courier Company (KG)'] = calcWeight(MergedTable.loc[i, 'Charged Weight'], MergedTable.loc[i, 'Zone'])

#Converting Zone to lowercase in courierRates for uniformity
courierRates['Zone'] = courierRates['Zone'].apply(lambda x: x.lower())

#Calculating ChargesAccX
def calcCharges(appWeight, shipmentType, zone):
    charges = 0
    weightSlab = weightslabs[zone]
    if shipmentType == 'Forward charges':
        fwdFixed = courierRates[courierRates['Zone'] == zone]['Forward Fixed Charge'].values[0]
        additional  = ((appWeight/weightSlab)-1) * courierRates[courierRates['Zone'] == zone]['Forward Additional Weight Slab Charge'].values[0]
        charges = fwdFixed + additional
    if shipmentType == 'Forward and RTO charges':
        fwdFixed = courierRates[courierRates['Zone'] == zone]['Forward Fixed Charge'].values[0]
        rtofixed = courierRates[courierRates['Zone'] == zone]['RTO Fixed Charge'].values[0]
        fwdadditional =  ((appWeight/weightSlab)-1) * courierRates[courierRates['Zone'] == zone]['Forward Additional Weight Slab Charge'].values[0]
        rtoadditional = ((appWeight/weightSlab)-1) * courierRates[courierRates['Zone'] == zone]['RTO Additional Weight Slab Charge'].values[0]
        charges = fwdFixed + rtofixed + fwdadditional + rtoadditional
    charges = round(charges, 2)
    return charges


for i in range(len(MergedTable)):
    MergedTable.loc[i, 'ChargesAccX'] = calcCharges(MergedTable.loc[i,'WeightbyX'], MergedTable.loc[i, 'Type of Shipment'], MergedTable.loc[i, 'ZonebyX'])


MergedTable['Charges Difference'] = MergedTable.apply(lambda x: x['ChargesAccX'] - x['Billing Amount (Rs.)'], axis=1)

print(MergedTable.info())

MergedTable.rename(columns={
    'AWB Code': 'AWB Number',
    'Charged Weight': 'Total weight as per Courier Company (KG)',
    'Total Order Weight (kg)' : 'Total weight as per X (KG)',
    'WeightbyX': 'Weight slab as per X (KG)',
    'ZonebyX': 'Delivery Zone as per X',
    'Zone': 'Delivery Zone charged by Courier Company',
    'ChargesAccX': 'Expected Charge as per X (Rs.)',
    'Billing Amount (Rs.)': 'Charges Billed by Courier Company (Rs.)',
    'Charges Difference': 'Difference Between Expected Charges and Billed Charges (Rs.)'
}, inplace=True)

cols = ['Order ID', 'AWB Number', 'Total weight as per X (KG)', 'Weight slab as per X (KG)',
        'Total weight as per Courier Company (KG)', 'Weight slab charged by Courier Company (KG)',
        'Delivery Zone as per X', 'Delivery Zone charged by Courier Company', 'Expected Charge as per X (Rs.)',
        'Charges Billed by Courier Company (Rs.)', 'Difference Between Expected Charges and Billed Charges (Rs.)']

Output = MergedTable[cols]

Output.to_excel('Results.xlsx')