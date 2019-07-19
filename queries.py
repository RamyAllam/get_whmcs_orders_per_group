orders_by_group_query_sql = "SELECT " \
                        "tblorders.id," \
                        "tblorders.ordernum," \
                        "tblorders.date," \
                        "tblorders.invoiceid," \
                        "tblorders.status," \
                        "tblorders.amount," \
                        "tblclients.id," \
                        "tblclients.firstname," \
                        "tblclients.lastname," \
                        "tblclients.companyname," \
                        "tblclients.email " \
                        "FROM tblorders " \
                        "INNER JOIN tblclients " \
                        "ON tblclients.id=tblorders.userid " \
                        "WHERE (tblorders.date LIKE '{}-%') " \
                        "AND (tblorders.status='Active') " \
                        "AND (tblclients.groupid={})"


# Query without passing client group ID
orders_query_sql = "SELECT " \
                        "tblorders.id," \
                        "tblorders.ordernum," \
                        "tblorders.date," \
                        "tblorders.invoiceid," \
                        "tblorders.status," \
                        "tblorders.amount," \
                        "tblclients.id," \
                        "tblclients.firstname," \
                        "tblclients.lastname," \
                        "tblclients.companyname," \
                        "tblclients.email " \
                        "FROM tblorders " \
                        "INNER JOIN tblclients " \
                        "ON tblclients.id=tblorders.userid " \
                        "WHERE (tblorders.date LIKE '{}-%') " \
                        "AND (tblorders.status='Active')"


# Query product status assigned for each order
product_status_query_sql = "SELECT id," \
                           " domain," \
                           " domainstatus," \
                           " orderid" \
                           " FROM tblhosting" \
                           " WHERE orderid={}"
