import pandas as pd
import numpy as np
import sys

def preprocess(sales, schemes, true_up):
	sales = pd.read_csv(sales)
	sales_prexo = sales[(sales['alpha_flag'] == 1) & ((sales['status'] == 'DELIVERED') | (sales['status'] == 'READY_TO_SHIP')) & (sales['prexo_bumpup_adjustment_amount'] != 0)]
	sales_prexo['order_date'] = pd.to_datetime(sales_prexo['order_date'].astype(int), format='%Y%m%d')
	sales_prexo['imei1'] = np.where(sales_prexo['imei1'].isnull(), '', 'IMEI' + sales_prexo['imei1'].astype(str))
	sales_prexo['imei1'] = sales_prexo['imei1'].str.replace(r'\.0$', '')

	prexo = pd.read_excel(schemes, sheet_name='Prexo')
	prexo.columns = prexo.columns.str.upper()
	prexo = prexo[['CLAIM ID', 'BRAND', 'FSN', 'TITLE', 'START DATE', 'END DATE', 'AMOUNT']]
	prexo = prexo.sort_values(by=['START DATE'], ascending=False)
	prexo['SOLD UNITS'] = 0
	prexo['BRAND SUPPORT'] = 0
	prexo['CLAIMED UNITS'] = 0
	prexo['AMOUNT CLAIMED'] = 0
	prexo['EXCESS(SHORT) UNITS'] = 0
	prexo['EXCESS(SHORT) AMOUNT'] = 0
	prexo['BRAND'] = np.where((prexo['BRAND'] == 'Redmi') | (prexo['BRAND'] == 'Xiaomi'), 'Redmi-Xiaomi', prexo['BRAND'])
	prexo['BRAND'] = prexo['BRAND'].str.upper()

	true_up = pd.read_excel(true_up, skiprows=1, sheet_name='Prexo')
	true_up.columns = true_up.columns.str.upper()
	true_up = true_up[['CLAIM ID', 'BRAND', 'ORDER EXTERNAL ID', 'IMEI1', 'EXPECTED CN']]
	true_up['CLAIM ID'] = true_up['CLAIM ID'].astype(str)
	true_up['IMEI1'] = 'IMEI' + true_up['IMEI1'].astype(str)
	true_up['IMEI1'] = true_up['IMEI1'].str.replace(r'\.0$', '')
	true_up['BRAND'] = true_up['BRAND'].str.upper()

	return sales_prexo, prexo, true_up

def inner_loop(sales_prexo, prexo_extract, true_up, i, df_brand):
	claim_id = prexo_extract['CLAIM ID'].iloc[i]
	fsn = prexo_extract['FSN'].iloc[i]
	start_date = prexo_extract['START DATE'].iloc[i]
	end_date = prexo_extract['END DATE'].iloc[i]
	brand_support = prexo_extract['AMOUNT'].iloc[i]

	df_sales = sales_prexo[(sales_prexo['product_id'] == fsn) & (sales_prexo['order_date'] >= start_date) & (sales_prexo['order_date'] <= end_date)]

	sales_prexo = sales_prexo[~sales_prexo.isin(df_sales)].dropna(how='all')

	df_sales['claim_id'] = claim_id
	df_sales = df_sales[['claim_id', 'order_external_id', 'order_date', 'product_id', 'product_title', 'status', 'imei1']]
	df_sales['amount'] = brand_support

	df_sales = df_sales.merge(true_up, left_on='imei1', right_on='IMEI1', how='left', indicator='claimed_trueup')
	df_sales['claimed_trueup'] = np.where(df_sales['claimed_trueup'] == 'both', 'Yes', 'No')
	df_sales = df_sales.rename(columns={'EXPECTED CN': 'amount_claimed'})
	df_sales['amount_claimed'] = df_sales['amount_claimed'].fillna(0)
	df_sales['excess(short)_claimed'] = df_sales['amount_claimed'] - df_sales['amount']
	df_brand = df_brand.append(df_sales[['claim_id', 'order_external_id', 'order_date', 'product_id', 'product_title', 'status', 'imei1', 'amount', 'claimed_trueup', 'amount_claimed', 'excess(short)_claimed']])
	# print(df_brand.info())
	prexo_extract['SOLD UNITS'].iloc[i] = df_sales['order_external_id'].count()
	prexo_extract['BRAND SUPPORT'].iloc[i] = df_sales['amount'].sum()
	prexo_extract['CLAIMED UNITS'].iloc[i] = df_sales['ORDER EXTERNAL ID'].count()
	prexo_extract['AMOUNT CLAIMED'].iloc[i] = df_sales['amount_claimed'].sum()
	prexo_extract['EXCESS(SHORT) UNITS'].iloc[i] = prexo_extract['CLAIMED UNITS'].iloc[i] - prexo_extract['SOLD UNITS'].iloc[i]
	prexo_extract['EXCESS(SHORT) AMOUNT'].iloc[i] = prexo_extract['AMOUNT CLAIMED'].iloc[i] - prexo_extract['BRAND SUPPORT'].iloc[i]

	return claim_id, fsn, df_brand, prexo_extract, sales_prexo

def write_to_excel(writer, summary_brand):
	sold_units = summary_brand['SOLD UNITS'].sum()
	brand_support = summary_brand['BRAND SUPPORT'].sum()
	claimed_units = summary_brand['CLAIMED UNITS'].sum()
	amount_claimed = summary_brand['AMOUNT CLAIMED'].sum()
	diff_units = summary_brand['EXCESS(SHORT) UNITS'].sum()
	diff_amount = summary_brand['EXCESS(SHORT) AMOUNT'].sum()
	summary_brand.loc[len(summary_brand)] = ['', '', '', 'TOTAL', '', '', '', sold_units, brand_support, claimed_units, amount_claimed, diff_units, diff_amount]

	summary_brand.to_excel(writer, sheet_name='Summary-Prexo', index=False)

	workbook = writer.book
	number_format = workbook.add_format({'num_format': '#,##0'})
	date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
	center_format = workbook.add_format()
	center_format.set_align('center')

	sheets = writer.sheets
	for sheet in sheets:
		if sheet == 'Summary-Prexo':
			writer.sheets[sheet].set_column('A:C', 25, None)
			writer.sheets[sheet].set_column('D:D', 65, None)
			writer.sheets[sheet].set_column('E:F', 25, date_format)
			writer.sheets[sheet].set_column('G:M', 25, number_format)
			
		else:
			writer.sheets[sheet].set_column('A:B', 25, None)
			writer.sheets[sheet].set_column('C:C', 25, date_format)
			writer.sheets[sheet].set_column('D:D', 25, None)
			writer.sheets[sheet].set_column('E:E', 65, None)
			writer.sheets[sheet].set_column('F:G', 25, None)
			writer.sheets[sheet].set_column('H:H', 25, number_format)
			writer.sheets[sheet].set_column('I:I', 25, center_format)
			writer.sheets[sheet].set_column('J:K', 25, number_format)

	writer.save()

	return

def main(sales, schemes, true_up):
	writer = pd.ExcelWriter('temp/Prexo Reco Test.xlsx')
	sales_prexo, prexo, true_up = preprocess(sales, schemes, true_up)
	brand_list = sorted(list(set(prexo['BRAND'].to_list())))
	summary_brand = pd.DataFrame()
	for brand in brand_list:
		df_brand = pd.DataFrame()
		prexo_extract = prexo[prexo['BRAND'] == brand]
		
		for i in range(len(prexo_extract)):
			claim_id, fsn, df_brand, prexo_extract, sales_prexo = inner_loop(sales_prexo, prexo_extract, true_up, i, df_brand)
			print(brand, i, claim_id, fsn, len(df_brand))

		df_brand = df_brand.sort_values(by=['claim_id', 'product_id'])

		if len(df_brand) > 0:
			df_brand.to_excel(writer, sheet_name=brand + '-Prexo', index=False)

		summary_brand = summary_brand.append(prexo_extract)
		# break
	summary_brand = summary_brand.sort_values(by=['BRAND', 'CLAIM ID', 'FSN'])
	write_to_excel(writer, summary_brand)

	return

# main("Clean Pricing+Prexo Sales File.csv", "Consolidated Portfolios June'23 Scheme V1.xlsx", "June 23 TU.xlsb")