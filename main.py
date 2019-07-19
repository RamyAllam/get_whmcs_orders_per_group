from modules import generate_report, workbook

'''
group_id_query: WHMCS client's GROUP ID to query (Optional)
product_status_query: Query Products assigned to those orders (Optional)
order_date_query: Year, and month for the orders to query (Required)
'''

# Uncomment any of the following
# generate_report(order_date_query="2019-06", group_id_query=10, product_status_query=False)
# generate_report(order_date_query="2019-06", group_id_query=10, product_status_query=True)
# generate_report(order_date_query="2019-06", product_status_query=True)
# generate_report(order_date_query="2019-06", product_status_query=False)

workbook.close()
