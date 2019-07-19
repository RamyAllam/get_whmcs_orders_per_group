import pymysql
import xlsxwriter
from queries import orders_by_group_query_sql, orders_query_sql, product_status_query_sql
from vars import *

# Open XLSX file and adding sheets
workbook = xlsxwriter.Workbook(export_file)

# Formatting CELLS
cells_titles_format = workbook.add_format({'bold': True})
cells_titles_format.set_bg_color(cell_background)


def get_whmcs_data(order_date_query, group_id_query=None, product_status_query=False):
    whmcs_data = {}

    try:
        # Open database connection
        db = pymysql.connect(db_host, db_user, db_pass, db_name)

        # prepare a cursor object using cursor() method
        cursor = db.cursor()

        # execute SQL query using execute() method.
        # If Orders query is based on group ID
        if group_id_query:
            cursor.execute(orders_by_group_query_sql.format(order_date_query, group_id_query))
        # Query all orders
        else:
            cursor.execute(orders_query_sql.format(order_date_query))

        # Fetch all the rows in a list of lists.
        results = cursor.fetchall()

        product_orderid_list = []
        for row in results:
            order_id = row[0]
            ordernum = row[1]
            order_date = row[2]
            order_invoiceid = row[3]
            order_status = row[4]
            order_amount = row[5]
            client_id = row[6]
            client_firstname = row[7]
            client_lastname = row[8]
            client_companyname = row[9]
            client_email = row[10]
            multi_products = False

            whmcs_data[row[0]] = {
                'order_id': str(order_id),
                'ordernum': str(ordernum),
                'order_date': str(order_date),
                'order_invoiceid': str(order_invoiceid),
                'order_status': str(order_status),
                'order_amount': str(order_amount),
                'client_id': str(client_id),
                'client_firstname': str(client_firstname),
                'client_lastname': str(client_lastname),
                'client_companyname': str(client_companyname),
                'client_email': str(client_email),
                'multi_products': multi_products
            }

            # If product_status_query is set to True, Query all the products assigned to this order
            if product_status_query:
                cursor.execute(product_status_query_sql.format(order_id))
                product_status_results = cursor.fetchall()

                # Check if results are returned or the product is deleted
                if product_status_results:
                    for product_status_row in product_status_results:
                        product_id = product_status_row[0]
                        product_domain = product_status_row[1]
                        product_domainstatus = product_status_row[2]
                        product_orderid = product_status_row[3]

                        # Check if the product ID is found in the product_orderid_list
                        # To verify if the order has multiple products or not
                        if product_orderid in product_orderid_list:
                            product_id = "Multiple Product IDs, please check manually"
                            product_domain = "Multiple Product Domains, please check manually"
                            product_domainstatus = "Multiple Status, please check manually"
                            multi_products = True

                        # Save the product ID in the list to search it later in the above condition
                        # to make sure the order has multiple products or not
                        product_orderid_list.append(product_orderid)

                        whmcs_data[row[0]].update(
                            {
                                'product_id': product_id,
                                'product_domain': product_domain,
                                'product_domainstatus': product_domainstatus,
                                'multi_products': multi_products
                            }
                        )
                        print("============================")
                else:
                    print("Not found {}".format(product_status_results))
                    whmcs_data[row[0]].update(
                        {
                            'product_id': None,
                            'product_domain': None,
                            'product_domainstatus': None
                        }
                    )

        return whmcs_data

    except Exception as e:
        print("Exception: {}".format(e))


def generate_report(order_date_query, group_id_query=None, product_status_query=False, row_count=1):
    orders_sheet_created = False

    worksheet_orders_details = ''

    if not orders_sheet_created:
        worksheet_orders_details = workbook.add_worksheet("WHMCS")
        # Set width for column
        worksheet_orders_details.set_column(order_id_col, order_id_col, 10)
        worksheet_orders_details.set_column(ordernum_col, ordernum_col, 15)
        worksheet_orders_details.set_column(order_date_col, order_date_col, 25)
        worksheet_orders_details.set_column(order_invoiceid_col, order_invoiceid_col, 15)
        worksheet_orders_details.set_column(order_status_col, order_status_col, 10)
        worksheet_orders_details.set_column(order_amount_col, order_amount_col, 10)
        worksheet_orders_details.set_column(client_id_col, client_id_col, 10)
        worksheet_orders_details.set_column(client_firstname_col, client_firstname_col, 20)
        worksheet_orders_details.set_column(client_lastname_col, client_lastname_col, 20)
        worksheet_orders_details.set_column(client_companyname_col, client_companyname_col, 20)
        worksheet_orders_details.set_column(client_email_col, client_email_col, 30)
        worksheet_orders_details.set_column(whmcs_client_profile_link_col, whmcs_client_profile_link_col, 70)
        worksheet_orders_details.set_column(whmcs_order_link_col, whmcs_order_link_col, 70)

        # Format to write ( row, column, content, format )
        worksheet_orders_details.write(0, order_id_col, "Order ID", cells_titles_format)
        worksheet_orders_details.write(0, ordernum_col, "Order Number", cells_titles_format)
        worksheet_orders_details.write(0, order_date_col, "Order Date", cells_titles_format)
        worksheet_orders_details.write(0, order_invoiceid_col, "Order Invoice ID", cells_titles_format)
        worksheet_orders_details.write(0, order_status_col, "Order Status", cells_titles_format)
        worksheet_orders_details.write(0, order_amount_col, "Order Amount", cells_titles_format)
        worksheet_orders_details.write(0, client_id_col, "CLient ID", cells_titles_format)
        worksheet_orders_details.write(0, client_firstname_col, "First Name", cells_titles_format)
        worksheet_orders_details.write(0, client_lastname_col, "Last Name", cells_titles_format)
        worksheet_orders_details.write(0, client_companyname_col, "Company Name", cells_titles_format)
        worksheet_orders_details.write(0, client_email_col, "Email", cells_titles_format)
        worksheet_orders_details.write(0, whmcs_client_profile_link_col, "WHMCS Client Profile", cells_titles_format)
        worksheet_orders_details.write(0, whmcs_order_link_col, "WHMCS Order Link", cells_titles_format)

        # If the user asked to query the products status assigned for this order
        if product_status_query:
            worksheet_orders_details.set_column(product_id_col, product_id_col, 15)
            worksheet_orders_details.set_column(product_domain_col, product_domain_col, 35)
            worksheet_orders_details.set_column(product_domainstatus_col, product_domainstatus_col, 15)
            worksheet_orders_details.set_column(whmcs_product_link_col, whmcs_product_link_col, 70)
            worksheet_orders_details.write(0, product_id_col, "Product ID", cells_titles_format)
            worksheet_orders_details.write(0, product_domain_col, "Product Hostname", cells_titles_format)
            worksheet_orders_details.write(0, product_domainstatus_col, "Product Status", cells_titles_format)
            worksheet_orders_details.write(0, whmcs_product_link_col, "WHMCS Product Link", cells_titles_format)

    # Loop through get_whmcs_data and get the results
    for item in get_whmcs_data(str(order_date_query), group_id_query, product_status_query).items():
        order_id = item[1]['order_id']
        ordernum = item[1]['ordernum']
        order_date = item[1]['order_date']
        order_invoiceid = item[1]['order_invoiceid']
        order_status = item[1]['order_status']
        order_amount = item[1]['order_amount']
        client_id = item[1]['client_id']
        client_firstname = item[1]['client_firstname']
        client_lastname = item[1]['client_lastname']
        client_companyname = item[1]['client_companyname']
        client_email = item[1]['client_email']

        whmcs_client_profile = "{}/clientssummary.php?userid={}".format(whmcs_url, client_id)
        whmcs_order_link = "{}/orders.php?action=view&id={}".format(whmcs_url, order_id)

        # IF the user set product_status_query to TRUE, loop through it's results
        if product_status_query:
            product_id = item[1]['product_id']
            product_domain = item[1]['product_domain']
            product_domainstatus = item[1]['product_domainstatus']
            multi_products_status = item[1]['multi_products']

            # Check if the product really is existed or it's deleted
            if product_id:
                print("multi_products_status {} - Product ID {}".format(multi_products_status, product_id))
                # Check if the order does not have multiple products
                if multi_products_status is not True:
                    whmcs_product_link = "{}/clientsservices.php?userid={}&id={}".format(whmcs_url, client_id, product_id)
                # If it has multiple products, return multiple links
                else:
                    whmcs_product_link = "MULTIPLE LINKS"

            # Return not found if it's deleted
            else:
                product_id = "Product Not Found"
                product_domain = "Product Not Found"
                product_domainstatus = "Product Not Found"
                whmcs_product_link = "Product Not Found"

            worksheet_orders_details.write(row_count, product_id_col, product_id)
            worksheet_orders_details.write(row_count, product_domain_col, product_domain)
            worksheet_orders_details.write(row_count, product_domainstatus_col, product_domainstatus)
            worksheet_orders_details.write(row_count, whmcs_product_link_col, whmcs_product_link)

        worksheet_orders_details.write(row_count, order_id_col, order_id)
        worksheet_orders_details.write(row_count, ordernum_col, ordernum)
        worksheet_orders_details.write(row_count, order_date_col, order_date)
        worksheet_orders_details.write(row_count, order_invoiceid_col, order_invoiceid)
        worksheet_orders_details.write(row_count, order_status_col, order_status)
        worksheet_orders_details.write(row_count, order_amount_col, order_amount)
        worksheet_orders_details.write(row_count, client_id_col, client_id)
        worksheet_orders_details.write(row_count, client_firstname_col, client_firstname)
        worksheet_orders_details.write(row_count, client_lastname_col, client_lastname)
        worksheet_orders_details.write(row_count, client_companyname_col, client_companyname)
        worksheet_orders_details.write(row_count, client_email_col, client_email)
        worksheet_orders_details.write(row_count, whmcs_client_profile_link_col, whmcs_client_profile)
        worksheet_orders_details.write(row_count, whmcs_order_link_col, whmcs_order_link)

        row_count += 1
