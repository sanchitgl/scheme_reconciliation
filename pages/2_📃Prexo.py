import streamlit as st 
import pandas as pd
from prexo import preprocess, inner_loop, write_to_excel
import os
# import streamlit_authenticator as stauth
import pickle 
from pathlib import Path 
import yaml
from PIL import Image
import time
# import plotly.graph_objects as go
import base64
# from page_config import page_setup
# from login_page import login_status

st.set_page_config(layout="wide",initial_sidebar_state ="collapsed")

# page_setup()

state = st.session_state

# authentication_status = login_status()

# #authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "ship_recon", "admin")
# #with placeholder.container():
# space, login, space = st.columns([1,3,1])
# with login:
#     name, authentication_status, username = authenticator.login('Login', 'main')
# state.authentication_status

# if authentication_status == False:
#     space, login, space = st.columns([1,3,1])
#     with login:
#         st.error("Username/Password is incorrect")



# if authentication_status:
    #placeholder.empty()
    #authenticator.logout('Logout', 'sidebar')

    # time.sleep(0.1)
def landing_page():
    st.markdown('''
    <style>
    .css-9s5bis.edgvbvh3 {
    display: block;
    }
    </style>
    ''', unsafe_allow_html=True)
    #with title:
    # emp,title,emp = st.columns([2,2,2])
    # with title:
    if 'submit' not in state:
        state.submit= False
    if 'response' not in state:
        state.response = []
    st.markdown("<h2 style='text-align: center; padding:0'>Prexo</h2>", unsafe_allow_html=True)
    #st.write('###')
    sales, schemes, true_up, submit = file_upload_form()
    #print(warehouse_reports)
    # try:
    if submit:
        state.submit = True

        with st.spinner('Please wait'):
            try:
                delete_temp()
            except:
                print()

            if sales is not None and schemes is not None and true_up is not None:
                writer = pd.ExcelWriter('temp/Prexo Reco Test.xlsx')
                print("preprocessing...")
                sales_prexo, prexo, true_up = preprocess(sales, schemes, true_up)
                # sales_clean, clean_pricing, true_up = preprocess(sales, schemes, true_up)
                # sales_prexo.to_csv('temp/sales_prexo.csv')
                # prexo.to_csv('temp/clean_prexo.csv')
                # true_up.to_csv('temp/true_up_prexo.csv')
                # sales_prexo = pd.read_csv('temp/sales_prexo.csv', index_col = None) 
                # prexo = pd.read_csv('temp/clean_prexo.csv', index_col = None) 
                # true_up = pd.read_csv('temp/true_up_prexo.csv', index_col = None) 
                # sales_clean = sales_clean.head(1000)

                print("preprocessing done.")

                brand_list = sorted(list(set(prexo['BRAND'].to_list())))
                summary_brand = pd.DataFrame()
                with st.empty():
                    for brand in ['VIVO','OPPO']:
                        df_brand = pd.DataFrame()
                        prexo_extract = prexo[prexo['BRAND'] == brand]
                        
                        for i in range(len(prexo_extract)):
                            claim_id, fsn, df_brand, prexo_extract, sales_prexo = inner_loop(sales_prexo, prexo_extract, true_up, i, df_brand)

                            col1, col2, col3, col4, col5 = st.columns(5)

                            with col1:
                                st.write(f'<h6>{"Brand"}</h5>', unsafe_allow_html=True)
                                st.text(str(brand))

                            with col2:
                                st.write(f'<h6>{"Number"}</h5>', unsafe_allow_html=True)
                                # st.header('Number')
                                st.text(str(i))

                            with col3:
                                st.write(f'<h6>{"Claim ID"}</h5>', unsafe_allow_html=True)
                                # st.header('Claim ID')
                                st.text(str(claim_id))

                            with col4:
                                st.write(f'<h6>{"FSN"}</h5>', unsafe_allow_html=True)
                                # st.header('FSN')
                                st.text(str(fsn))
                            
                            with col5:
                                st.write(f'<h6>{"Brand Length"}</h5>', unsafe_allow_html=True)
                                # st.header('Brand Length')
                                st.text(str(len(df_brand)))
                            # st.write("brand: "+str(brand) + ", i: " + str(i) + ", claim_id: " + str(claim_id) + ", fsn: " + str(fsn) + ", Brand Length: " + str(len(df_brand)))

                        df_brand = df_brand.sort_values(by=['claim_id', 'product_id'])

                        if len(df_brand) > 0:
                            df_brand.to_excel(writer, sheet_name=brand + '-Prexo', index=False)

                        summary_brand = summary_brand.append(prexo_extract)
                        # break

                    summary_brand = summary_brand.sort_values(by=['BRAND', 'CLAIM ID', 'FSN'])
                    write_to_excel(writer, summary_brand)

        emp, but, empty = st.columns([2.05,1.2,1.5]) 
        with but:
            with open('temp/Prexo Reco Test.xlsx', 'rb') as my_file:
                click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'Prexo Reco Test.xlsx', 
                mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        if state.submit == True:
            if state.response != {}:
                emp, but, empty = st.columns([2.05,1.2,1.5]) 
                with but:
                    with open('temp/Prexo Reco Test.xlsx', 'rb') as my_file:
                        click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'Prexo Reco Test.xlsx', 
                        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    # except:
    #     st.error("Run failed, kindly check if the inputs are valid")

def delete_temp():
    os.remove("temp/Prexo Reco Test.xlsx")

def file_upload_form():
    colour = "#89CFF0"
    with st.form(key = 'ticker',clear_on_submit=False):
        text, upload = st.columns([2.5,3]) 
        with text:
            st.write("###")
            st.write("###")
            st.write(f'<h5>{"&nbsp; Upload Sales File:"}</h5>', unsafe_allow_html=True)
        with upload:
            shipment_instructions = st.file_uploader("",key = 'ship_ins', accept_multiple_files=False)

        text, upload = st.columns([2.5,3])
        with text:
            st.write("###")
            st.write("###")
            st.write(f'<h5>{"&nbsp; Upload Scheme File:"}<h5>', unsafe_allow_html=True)
        with upload:
            warehouse_reports = st.file_uploader("",key = 'ware_rep', accept_multiple_files=False)

        text, upload = st.columns([2.5,3])
        with text:
            st.write("###")
            st.write("###")
            st.write(f'<h5> {"&nbsp; Upload True Up File:"}<h5>', unsafe_allow_html=True)
        with upload:
            inventory_ledger = st.file_uploader("",key = 'inv_led')
        
        a,button,b = st.columns([2,1.2,1.5]) 
        with button:
            st.write('###')
            submit = st.form_submit_button(label = "Start Reconciliation")
            #submit = st.button(label="Start Reconciliation")
    return shipment_instructions, warehouse_reports, inventory_ledger, submit
    

    

landing_page()