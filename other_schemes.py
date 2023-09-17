import pandas as pd
import datetime
import sys
# import openpyxl as oxl
import numpy as np

def preprocess(sales, schemes, true_up):
	required = ['PBO', 'Prepaid', 'Coupon', 'Supercoins']

	sales = pd.read_csv(sales)
	sales = sales[(sales['is_alpha_seller'] == True) & ((sales['status'] == 'DELIVERED') | (sales['status'] == 'READY_TO_SHIP'))]
	sales['order_date'] = pd.to_datetime(sales['order_date'], format='%Y-%m-%d')
	sales['imei1'] = np.where(sales['imei1'].isnull(), '', 'IMEI' + sales['imei1'].astype(str))
	sales['imei1'] = sales['imei1'].str.replace(r'\.0$', '')

	schemes_file = pd.ExcelFile(schemes)
	schemes_sheets = schemes_file.sheet_names

	true_up_file = pd.ExcelFile(true_up)
	true_up_sheets = true_up_file.sheet_names

	schemes_sheets = [sheet for sheet in schemes_sheets if sheet in required]
	schemes_sheets = sorted(schemes_sheets)
	true_up_sheets = [sheet for sheet in true_up_sheets if sheet in required]

	schemes_dict = {}
	true_up_dict = {}

	for sheet in schemes_sheets:
		df_scheme = schemes_file.parse(sheet)
		df_scheme.columns = df_scheme.columns.str.upper()
		schemes_dict[sheet] = df_scheme

	for sheet in true_up_sheets:
		df_true_up = true_up_file.parse(sheet, skiprows=1)
		df_true_up.columns = df_true_up.columns.str.upper()
		true_up_dict[sheet] = df_true_up

	schemes_file.close()
	true_up_file.close()

	return sales, schemes_dict, true_up_dict, schemes_sheets, true_up_sheets

def preprocess_sheet(sheet, schemes_dict, true_up_dict, schemes_sheets, true_up_sheets):
	# for sheet in schemes_sheets:
		# print(sheet)
	scheme = schemes_dict[sheet]
	scheme = scheme[['CLAIM ID', 'BRAND', 'FSN', 'OFFER ID', 'AMOUNT']]
	scheme['CLAIM ID'] = scheme['CLAIM ID'].astype(str)
	scheme['BRAND'] = scheme['BRAND'].str.upper()
	scheme['OFFER ID'] = scheme['OFFER ID'].fillna('')

	scheme['SOLD UNITS'] = 0
	scheme['BRAND SUPPORT'] = 0
	scheme['CLAIMED UNITS'] = 0
	scheme['AMOUNT CLAIMED'] = 0
	scheme['EXCESS(SHORT) UNITS'] = 0
	scheme['EXCESS(SHORT) AMOUNT'] = 0

	if sheet in true_up_sheets:
		true_up = true_up_dict[sheet]
	
		if sheet == 'PBO':
			true_up = true_up.rename(columns={'IMEI_1': 'IMEI1', 'WITHOUT_PROPORTIONATE_SHARING': 'EXPECTED CN', 'CHECKOUT_ID': 'ORDER EXTERNAL ID'})
			true_up['IMEI1'] = true_up['IMEI1'].str.replace('"', '')

		true_up = true_up[['CLAIM ID', 'BRAND', 'ORDER EXTERNAL ID', 'IMEI1', 'EXPECTED CN']]
		true_up['CLAIM ID'] = true_up['CLAIM ID'].astype(str)
		true_up['IMEI1'] = true_up['IMEI1'].str.replace("'", "")
		true_up['IMEI1'] = 'IMEI' + true_up['IMEI1'].astype(str)
		true_up['IMEI1'] = true_up['IMEI1'].str.replace(r'\.0$', '')
		true_up['BRAND'] = true_up['BRAND'].str.upper()

	else:
		true_up = pd.DataFrame(columns=['CLAIM ID', 'BRAND', 'ORDER EXTERNAL ID', 'IMEI1', 'EXPECTED CN'])
	
	return scheme, true_up

def inner_loop(df_claim, sales, true_up, claim_id, brand_support, fsn, offer_id):

	df_sales = sales[(sales['fsn'] == fsn) & (sales['offer_id'] == offer_id)]
	df_sales['claim_id'] = claim_id
	df_sales = df_sales[['claim_id', 'fsn', 'offer_id', 'order_external_id', 'order_date', 'status', 'imei1']]
	df_sales['amount'] = brand_support
	df_sales = df_sales.merge(true_up, left_on='imei1', right_on='IMEI1', how='left', indicator='claimed_trueup')
	df_sales['claimed_trueup'] = np.where(df_sales['claimed_trueup'] == 'both', 'Yes', 'No')
	df_sales = df_sales.rename(columns={'EXPECTED CN': 'amount_claimed'})
	df_sales['amount_claimed'] = df_sales['amount_claimed'].fillna(0)
	df_sales['excess(short)_claimed'] = df_sales['amount_claimed'] - df_sales['amount']
	df_claim = df_claim.append(df_sales)
	
	return df_claim

def write_to_excel(writer, summary_brand, sheet):
	sold_units = summary_brand['SOLD UNITS'].sum()
	brand_support = summary_brand['BRAND SUPPORT'].sum()
	claimed_units = summary_brand['CLAIMED UNITS'].sum()
	amount_claimed = summary_brand['AMOUNT CLAIMED'].sum()
	diff_units = summary_brand['EXCESS(SHORT) UNITS'].sum()
	diff_amount = summary_brand['EXCESS(SHORT) AMOUNT'].sum()
	summary_brand.loc[len(summary_brand)] = ['', '', '', 'TOTAL', '', sold_units, brand_support, claimed_units, amount_claimed, diff_units, diff_amount]

	summary_brand.to_excel(writer, sheet_name='Summary-' + sheet, index=False)

	workbook = writer.book
	number_format = workbook.add_format({'num_format': '#,##0'})
	date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
	center_format = workbook.add_format()
	center_format.set_align('center')

	sheets = writer.sheets
	for sheet in sheets:
		if 'Summary' in sheet:
			writer.sheets[sheet].set_column('A:D', 25, None)
			writer.sheets[sheet].set_column('E:K', 25, number_format)
			
		else:
			writer.sheets[sheet].set_column('A:G', 25, None)
			writer.sheets[sheet].set_column('H:H', 25, number_format)
			writer.sheets[sheet].set_column('I:I', 25, center_format)
			writer.sheets[sheet].set_column('J:K', 25, number_format)
				
	# writer.save()

def main(sales, schemes, true_up):
	writer = pd.ExcelWriter('temp/Other Schemes Reco Test2.xlsx')
	sales, schemes_dict, true_up_dict, schemes_sheets, true_up_sheets = preprocess(sales, schemes, true_up)
	for sheet in schemes_sheets:
		scheme, true_up = preprocess_sheet(sheet, schemes_dict, true_up_dict, schemes_sheets, true_up_sheets)
		brand_list = sorted(list(set(scheme['BRAND'].to_list())))
		summary_brand = pd.DataFrame()
		for brand in brand_list:
			# print(brand)
			df_brand = pd.DataFrame()
			scheme_extract = scheme[scheme['BRAND'] == brand]
			for i in range(len(scheme_extract)):
				claim_id = scheme_extract['CLAIM ID'].iloc[i]
				fsn = scheme_extract['FSN'].iloc[i]
				brand_support = scheme_extract['AMOUNT'].iloc[i]
				offer = scheme_extract['OFFER ID'].iloc[i]
				offer_ids = offer.split(',')
				
				df_claim = pd.DataFrame()
				for offer_id in offer_ids:
					print(sheet, brand, i, claim_id, fsn, offer_id)
					df_claim = inner_loop(df_claim, sales, true_up, claim_id, brand_support, fsn, offer_id)
				df_brand = df_brand.append(df_claim)

				scheme_extract['SOLD UNITS'].iloc[i] = df_claim['order_external_id'].count()
				scheme_extract['BRAND SUPPORT'].iloc[i] = df_claim['amount'].sum()
				scheme_extract['CLAIMED UNITS'].iloc[i] = df_claim['ORDER EXTERNAL ID'].count()
				scheme_extract['AMOUNT CLAIMED'].iloc[i] = df_claim['amount_claimed'].sum()
				scheme_extract['EXCESS(SHORT) UNITS'].iloc[i] = scheme_extract['CLAIMED UNITS'].iloc[i] - scheme_extract['SOLD UNITS'].iloc[i]
				scheme_extract['EXCESS(SHORT) AMOUNT'].iloc[i] = scheme_extract['AMOUNT CLAIMED'].iloc[i] - scheme_extract['BRAND SUPPORT'].iloc[i]
				# break
			df_brand = df_brand.sort_values(by=['claim_id', 'fsn', 'offer_id', 'order_external_id'])

			df_brand_excel = df_brand[['claim_id', 'fsn', 'offer_id', 'order_external_id', 'order_date', 'status', 'imei1', 'amount', 'claimed_trueup', 'amount_claimed', 'excess(short)_claimed']]#.set_index(['claim_id', 'fsn', 'offer_id', 'order_external_id'])
			df_brand_excel.to_excel(writer, sheet_name=brand + '-' + sheet, index=False)

			# df_brand_grouped = df_brand.groupby(['claim_id', 'fsn', 'offer_id']).agg({'order_external_id': 'count', 'amount': 'sum', 'IMEI1': 'count', 'amount_claimed': 'sum', 'excess(short)_claimed': 'sum'})

			summary_brand = summary_brand.append(scheme_extract)
			# break
		summary_brand = summary_brand.sort_values(by=['BRAND', 'CLAIM ID', 'FSN'])
		write_to_excel(writer, summary_brand, sheet)

	writer.save()
	
	return

# main("Sales File IMEI+OFFER ID.csv", "Consolidated Portfolios Jul'23 Scheme V2- Copy.xlsx", "July'23TU_New_1.xlsb")
