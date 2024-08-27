import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# Function to clean cuit
def clean_cuit(cuit):
    if pd.isna(cuit):
        return ''
    return ''.join(filter(str.isdigit, str(cuit)))

# Function to split name into first name and last name
def split_name(name):
    if isinstance(name, str):
        parts = name.split(' ', 1)
        firstname = parts[0]
        lastname = parts[1] if len(parts) > 1 else ''
        return firstname, lastname
    return '', ''  # Return empty strings if name is not a valid string

# Function to create suffix from name
def create_suffix(name):
    if isinstance(name, str):
        return ''.join([word[0].upper() for word in name.split()])
    return ''  # Return empty string if name is not a valid string

# Function to process the uploaded files
def process_files(transactions_file, buyers_file, products_file=None, template_file=None):
    # Load data
    transactions_df = pd.read_excel(transactions_file)
    buyers_df = pd.read_excel(buyers_file)
    
    # Clean cuit column
    transactions_df['cuit'] = transactions_df['cuit'].apply(clean_cuit)

    # Merge transactions with buyers data
    transactions_df = transactions_df.merge(buyers_df[['Buyer Name', 'Buyer Email', 'Tax ID']],
                                             left_on='cuit', right_on='Tax ID', how='left')

    # Process transactions_df
    transactions_df['Buyer Name'] = transactions_df['Buyer Name'].fillna('')
    transactions_df['firstname'], transactions_df['lastname'] = zip(*transactions_df['Buyer Name'].apply(split_name))
    transactions_df['suffix'] = transactions_df['Buyer Name'].apply(create_suffix)
    transactions_df['Created_at'] = pd.to_datetime(transactions_df['fecha']) - pd.Timedelta(days=2)
    transactions_df['Created_at'] = transactions_df['Created_at'].dt.strftime('%d/%m/%Y')
    transactions_df['fecha'] = pd.to_datetime(transactions_df['fecha']).dt.strftime('%d/%m/%Y')

    # Rename and reorder columns for Transactions_Sample
    transactions_sample_df = transactions_df[['comprobante', 'fecha', 'Created_at', 'suffix', 'firstname', 'lastname',
                                              'Buyer Email', 'art_id', 'art_nombre', 'total', 'art_Cantidad', 'precio_unidad']]
    column_mapping = {
        'comprobante': 'Invoice#',
        'fecha': 'Date',
        'Created_at': 'Created_at',
        'suffix': 'Suffix',
        'firstname': 'First Name',
        'lastname': 'Last Name',
        'Buyer Email': 'Customer Email',
        'art_id': 'SKU',
        'art_nombre': 'Description',
        'total': 'Invoice total Inc',
        'art_Cantidad': 'Total Units',
        'precio_unidad': 'Unit Value'
    }
    transactions_sample_df = transactions_sample_df.rename(columns=column_mapping)
    expected_order = ['Invoice#', 'Date', 'Created_at', 'Suffix', 'First Name', 'Last Name',
                       'Customer Email', 'SKU', 'Description', 'Invoice total Inc', 'Total Units', 'Unit Value']
    transactions_sample_df = transactions_sample_df[expected_order]

    # Handle Products_Sample
    if products_file:
        try:
            products_df = pd.read_excel(products_file, sheet_name='Products_Sample')
        except ValueError:
            st.warning("Products_Sample sheet not found. Creating an empty Products_Sample.")
            products_df = pd.DataFrame(columns=[
                'sku', 'attribute_set_code', 'product_type', 'categories', 'category_ids',
                'product_websites', 'name', 'description', 'short_description', 'weight',
                'product_online', 'visibility', 'price', 'url_key', 'thumbnail_image',
                'small_image', 'base_image', 'swatch_image', 'qty', 'is_in_stock',
                'additional_attributes', 'seller_id', 'source_code'])
    else:
        st.warning("No Products file provided. Creating an empty Products_Sample.")
        products_df = pd.DataFrame(columns=[
            'sku', 'attribute_set_code', 'product_type', 'categories', 'category_ids',
            'product_websites', 'name', 'description', 'short_description', 'weight',
            'product_online', 'visibility', 'price', 'url_key', 'thumbnail_image',
            'small_image', 'base_image', 'swatch_image', 'qty', 'is_in_stock',
            'additional_attributes', 'seller_id', 'source_code'])

    # Load the template file
    if template_file:
        template_df = pd.read_excel(template_file)
    else:
        st.warning("No template file provided. Creating an empty Customer_Sample with the right columns.")
        template_df = pd.DataFrame(columns=[
            'email', '_website', '_store', 'confirmation', 'created_at', 'created_in',
            'disable_auto_group_change', 'dob', 'firstname', 'gender', 'group_id', 'lastname',
            'middlename', 'password_hash', 'prefix', 'rp_token', 'rp_token_created_at', 'store_id',
            'suffix', 'taxvat', 'cnpj', 'website_id', 'password', '_address_city', '_address_company',
            '_address_country_id', '_address_fax', '_address_firstname', '_address_lastname',
            '_address_middlename', '_address_postcode', '_address_prefix', '_address_region',
            '_address_street', '_address_suffix', '_address_telephone', '_address_vat_id',
            '_address_default_billing_', '_address_default_shipping_'])
    
    # Create Customer_Sample DataFrame
    customer_sample_df = template_df.copy()
    customer_sample_df['email'] = transactions_df['Buyer Email']
    customer_sample_df['_website'] = 'Argentina'
    customer_sample_df['_store'] = 'RM_ARG_VW'
    customer_sample_df['created_in'] = 'WIX'
    customer_sample_df['disable_auto_group_change'] = 0
    customer_sample_df['firstname'] = transactions_df['firstname']
    customer_sample_df['lastname'] = transactions_df['lastname']
    customer_sample_df['prefix'] = 'NHE'
    customer_sample_df['suffix'] = transactions_df['suffix']
    customer_sample_df['taxvat'] = transactions_df['Tax ID']
    customer_sample_df['website_id'] = 2
    customer_sample_df['_address_city'] = buyers_df['Shipping City']
    customer_sample_df['_address_fax'] = 'ARG'
    customer_sample_df['_address_firstname'] = transactions_df['firstname']
    customer_sample_df['_address_lastname'] = transactions_df['lastname']
    customer_sample_df['_address_postcode'] = buyers_df['Shipping Address'].str.extract(r'(\d+),AR')[0]
    customer_sample_df['_address_region'] = buyers_df['Shipping City']
    customer_sample_df['_address_street'] = buyers_df['Shipping Address'].str.extract(r'(.*?),\d+')[0]
    customer_sample_df['_address_telephone'] = buyers_df['Buyer Phone Number']
    customer_sample_df['_address_default_billing_'] = 1
    customer_sample_df['_address_default_shipping_'] = 1

   
    # Prepare output for download
    today = datetime.now().strftime('%d%m%Y')
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        customer_sample_df.to_excel(writer, sheet_name='Customer_Sample', index=False)
        products_df.to_excel(writer, sheet_name='Products_Sample', index=False)
        transactions_sample_df.to_excel(writer, sheet_name='Transactions_Sample', index=False)
    output.seek(0)  # Rewind the BytesIO object

    return output, f'BULK_ORDER_{today}.xlsx'

# Streamlit app
st.title('Bulk Order Processing')

# Upload files
uploaded_buyers_file = st.file_uploader("Upload Buyers Info Excel File", type=["xlsx"])
uploaded_transactions_file = st.file_uploader("Upload Transactions Excel File", type=["xlsx"])
uploaded_products_file = st.file_uploader("Upload Products Excel File", type=["xlsx"], help="Optional: You can skip this file if not needed.")
uploaded_template_file = st.file_uploader("Upload Bulk Template Excel File", type=["xlsx"])

if uploaded_buyers_file and uploaded_transactions_file and uploaded_template_file:
    with st.spinner('Processing files...'):
        output, file_name = process_files(uploaded_transactions_file, uploaded_buyers_file, uploaded_products_file, uploaded_template_file)
        
        # Provide download button for the result
        st.download_button(
            label="Download BULK ORDER File",
            data=output,
            file_name=file_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.warning('Please upload all required files (Buyers Info, Transactions, and Bulk Template). Products file is optional.')
