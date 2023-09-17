import pandas as pd
import numpy as np
import sys
import datetime

def preprocess(sales, schemes, true_up):
	sales = pd.read_csv(sales)
	# sales = sales.head(1000)
	sales['order_date'] = pd.to_datetime(sales['order_date'], format='%Y%m%d')
	sales_clean = sales[(sales['alpha_flag'] == 1) & ((sales['status'] == 'DELIVERED') | (sales['status'] == 'READY_TO_SHIP'))]
	sales_clean['imei1'] = np.where(sales_clean['imei1'].isnull(), '', 'IMEI' + sales_clean['imei1'].astype(str))
	sales_clean['imei1'] = sales_clean['imei1'].str.replace(r'\.0$', '')

	clean_pricing = pd.read_excel(schemes, sheet_name='Clean Pricing')
	# clean_pricing = clean_pricing.head(1000)
	clean_pricing.columns = clean_pricing.columns.str.upper()
	clean_pricing = clean_pricing[['CLAIM ID', 'BRAND', 'FSN', 'TITLE', 'LISTING PRICE', 'START DATE', 'END DATE', 'AMOUNT']]
	clean_pricing['START DATE'] = pd.to_datetime(clean_pricing['START DATE'])
	clean_pricing['END DATE'] = pd.to_datetime(clean_pricing['END DATE'])
	clean_pricing = clean_pricing.sort_values(by=['START DATE'], ascending=False)
	clean_pricing['SOLD UNITS'] = 0
	clean_pricing['BRAND SUPPORT'] = 0
	clean_pricing['CLAIMED UNITS'] = 0
	clean_pricing['AMOUNT CLAIMED'] = 0
	clean_pricing['EXCESS(SHORT) UNITS'] = 0
	clean_pricing['EXCESS(SHORT) AMOUNT'] = 0
	clean_pricing['BRAND'] = np.where((clean_pricing['BRAND'] == 'Redmi') | (clean_pricing['BRAND'] == 'Xiaomi'), 'Redmi-Xiaomi', clean_pricing['BRAND'])
	clean_pricing['BRAND'] = clean_pricing['BRAND'].str.upper()

	true_up_file = pd.ExcelFile(true_up)
	true_up_sheets = true_up_file.sheet_names
	true_up_sheets = [sheet for sheet in true_up_sheets if sheet[:2] == 'CP']

	true_up = pd.DataFrame()
	for true_up_sheet in true_up_sheets:
		df_true_up = true_up_file.parse(true_up_sheet, skiprows=1)
		true_up = true_up.append(df_true_up)

	true_up.columns = true_up.columns.str.upper()
	true_up = true_up[['CLAIM ID', 'BRAND', 'ORDER EXTERNAL ID', 'IMEI1', 'EXPECTED CN']]
	true_up['CLAIM ID'] = true_up['CLAIM ID'].astype(str)
	true_up['IMEI1'] = 'IMEI' + true_up['IMEI1'].astype(str)
	true_up['IMEI1'] = true_up['IMEI1'].str.replace(r'\.0$', '')
	true_up['BRAND'] = true_up['BRAND'].str.upper()

	
	return sales_clean, clean_pricing, true_up

def inner_loop(sales_clean, clean_extract, true_up, i, df_brand):
	claim_id = clean_extract['CLAIM ID'].iloc[i]
	fsn = clean_extract['FSN'].iloc[i]
	start_date = clean_extract['START DATE'].iloc[i]
	end_date = clean_extract['END DATE'].iloc[i]
	max_price = clean_extract['LISTING PRICE'].iloc[i]
	brand_support = clean_extract['AMOUNT'].iloc[i]

	df_sales = sales_clean[(sales_clean['product_id'] == fsn) & (sales_clean['order_date'] >= start_date) & (sales_clean['order_date'] <= end_date) & (sales_clean['listing_price'] <= max_price)]

	sales_clean = sales_clean[~sales_clean.isin(df_sales)].dropna(how='all')

	df_sales['claim_id'] = claim_id
	df_sales = df_sales[['claim_id', 'order_external_id', 'order_date', 'product_id', 'product_title', 'listing_price', 'status', 'imei1']]
	df_sales['amount'] = brand_support

	df_sales = df_sales.merge(true_up, left_on='imei1', right_on='IMEI1', how='left', indicator='claimed_trueup')
	df_sales['claimed_trueup'] = np.where(df_sales['claimed_trueup'] == 'both', 'Yes', 'No')
	df_sales = df_sales.rename(columns={'EXPECTED CN': 'amount_claimed'})
	df_sales['amount_claimed'] = df_sales['amount_claimed'].fillna(0)
	df_sales['excess(short)_claimed'] = df_sales['amount_claimed'] - df_sales['amount']

	df_brand = df_brand.append(df_sales[['claim_id', 'order_external_id', 'order_date', 'product_id', 'product_title', 'listing_price', 'status', 'imei1', 'amount', 'claimed_trueup', 'amount_claimed', 'excess(short)_claimed']])

	clean_extract['SOLD UNITS'].iloc[i] = df_sales['order_external_id'].count()
	clean_extract['BRAND SUPPORT'].iloc[i] = df_sales['amount'].sum()
	clean_extract['CLAIMED UNITS'].iloc[i] = df_sales['ORDER EXTERNAL ID'].count()
	clean_extract['AMOUNT CLAIMED'].iloc[i] = df_sales['amount_claimed'].sum()
	clean_extract['EXCESS(SHORT) UNITS'].iloc[i] = clean_extract['CLAIMED UNITS'].iloc[i] - clean_extract['SOLD UNITS'].iloc[i]
	clean_extract['EXCESS(SHORT) AMOUNT'].iloc[i] = clean_extract['AMOUNT CLAIMED'].iloc[i] - clean_extract['BRAND SUPPORT'].iloc[i]

	return claim_id, fsn, df_brand, clean_extract, sales_clean

def write_to_excel(writer, summary_brand):
	sold_units = summary_brand['SOLD UNITS'].sum()
	brand_support = summary_brand['BRAND SUPPORT'].sum()
	claimed_units = summary_brand['CLAIMED UNITS'].sum()
	amount_claimed = summary_brand['AMOUNT CLAIMED'].sum()
	diff_units = summary_brand['EXCESS(SHORT) UNITS'].sum()
	diff_amount = summary_brand['EXCESS(SHORT) AMOUNT'].sum()
	summary_brand.loc[len(summary_brand)] = ['', '', '', 'TOTAL', '', '', '', '', sold_units, brand_support, claimed_units, amount_claimed, diff_units, diff_amount]

	summary_brand.to_excel(writer, sheet_name='Summary-CleanPricing', index=False)

	workbook = writer.book
	number_format = workbook.add_format({'num_format': '#,##0'})
	date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
	center_format = workbook.add_format()
	center_format.set_align('center')

	sheets = writer.sheets
	for sheet in sheets:
		if sheet == 'Summary-CleanPricing':
			writer.sheets[sheet].set_column('A:C', 25, None)
			writer.sheets[sheet].set_column('D:D', 65, None)
			writer.sheets[sheet].set_column('E:E', 25, number_format)
			writer.sheets[sheet].set_column('F:G', 25, date_format)
			writer.sheets[sheet].set_column('H:N', 25, number_format)

		else:
			writer.sheets[sheet].set_column('A:B', 25, None)
			writer.sheets[sheet].set_column('C:C', 25, date_format)
			writer.sheets[sheet].set_column('D:D', 25, None)
			writer.sheets[sheet].set_column('E:E', 65, None)
			writer.sheets[sheet].set_column('F:F', 25, number_format)
			writer.sheets[sheet].set_column('G:H', 25, None)
			writer.sheets[sheet].set_column('I:I', 25, number_format)
			writer.sheets[sheet].set_column('J:J', 25, center_format)
			writer.sheets[sheet].set_column('K:L', 25, number_format)

	writer.save()

	return

# def main(sales, schemes, true_up):
# 	writer = pd.ExcelWriter('temp/Clean Pricing Test2.xlsx')
# 	sales_clean, clean_pricing, true_up = preprocess(sales, schemes, true_up)
# 	brand_list = sorted(list(set(clean_pricing['BRAND'].to_list())))
# 	summary_brand = pd.DataFrame()
# 	for brand in brand_list:
# 		df_brand = pd.DataFrame()
# 		clean_extract = clean_pricing[clean_pricing['BRAND'] == brand]
		
# 		for i in range(len(clean_extract)):
# 			claim_id, fsn, df_brand, clean_extract, sales_clean = inner_loop(sales_clean, clean_extract, true_up, i, df_brand)
# 			print(brand, i, claim_id, fsn)

# 		df_brand = df_brand.sort_values(by=['claim_id', 'product_id'])

# 		if len(df_brand) > 0:
# 			df_brand.to_excel(writer, sheet_name=brand + '-CleanPricing', index=False)

# 		summary_brand = summary_brand.append(clean_extract)
# 		break
# 	summary_brand = summary_brand.sort_values(by=['BRAND', 'CLAIM ID', 'FSN'])
# 	write_to_excel(writer, summary_brand)

# 	return

# main("Clean Pricing+Prexo Sales File.csv", "Consolidated Portfolios June'23 Scheme V1.xlsx", "June 23 TU.xlsb")
 