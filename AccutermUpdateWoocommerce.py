# from textwrap import indent
# from typing import ValuesView
from woocommerce import API
# import json
import csv
import mysql.connector
from openpyxl import load_workbook

#=================================================================================

wcapi = API(
  url="https://url.com",
  consumer_key="consumer_key",
  consumer_secret="consumer_secret",
  timeout=50
)

conn = mysql.connector.connect(
  database='db',
  host='host',
  password='password',
  user='user'
)

#=========================== HELPER FUNCTIONS =============================================

#Prints mySQL.connector error
def print_error(error):
    print(error)
    print("Error Code:", error.errno)
    print("SQLSTATE", error.sqlstate)
    print("Message", error.msg)

#Drop all existing tables
def drop_tables(cursor):
    cursor.execute('SHOW TABLES;')
    for table in cursor.fetchall(): 
      try:
        print(f"Deleting `{table[0]}` ...")
        cursor.execute(f"DROP TABLE {table[0]};")
      except mysql.connector.Error as error:  print_error(error)

#Returns SQL-Formatted CREATE TABLE statement as a string key value
def table_constructor(table_name):
    TABLES = {}

    TABLES = dict.fromkeys(['valid_items', 'raw_items'], f"""
    CREATE TABLE {table_name} (
      product_id VARCHAR(11) NOT NULL PRIMARY KEY,
      product_description VARCHAR(100),
      description2 VARCHAR(100),
      description3 VARCHAR(100),
      reg_price INT(11),
      pu_price INT(11),
      avail INT(11),
      is_continuing BOOLEAN NOT NULL
      )""")

    TABLES['website_items'] = ("""
      CREATE TABLE website_items (
        id INT NOT NULL PRIMARY KEY,
        name VARCHAR(11) NOT NULL,
        in_stock BOOLEAN NOT NULL,
        reg_price INT(11) DEFAULT 0
      )""" )

    return TABLES[table_name]

#Return list of column names in 'table_name
def get_table_columns(cursor, table_name):
    try:
      cursor.execute(f"SELECT * FROM {table_name}")
      field_names = [i[0] for i in cursor.description]
      return field_names

    except mysql.connector.Error as error:
      print_error(error)

#Creates mySQL 'table_name' from a pre-written SQL-formatted dictionary in 'table_constructor'
def create_table(cursor, table_name):
    try:
      cursor.execute(f"{table_constructor(table_name)};")

    except mysql.connector.Error as error:
      print_error(error)

#Inserts 'values' into any existing 'table_name'
def insert_table(cursor, table_name, values):
    try:

      #Stores list of column names into string from table_name
      table_columns = f"{get_table_columns(cursor, table_name)}"
      text_table = table_columns.maketrans('', '', '[]\'')
      table_columns = table_columns.translate(text_table)

      #Insert Data into Table
      cursor.execute(f"""
        INSERT INTO {table_name}
        ({table_columns})
        VALUES {values}
        """)

    except mysql.connector.IntegrityError as error: pass
    except mysql.connector.Error as error:  print_error(error)

#Returns SQL Select statement with conditions to filter out invalid items
def sql_get_valid_items(table_name):
    return f'''
    SELECT *
    FROM {table_name}
    WHERE product_id NOT LIKE '%-SR'
      AND product_id Not Like '%-DRM'
      AND product_id Not Like '%-DR'
      AND product_id Not Like '%-C'
      AND product_id Not Like '%-M'
      AND product_id Not Like '%-N'
      AND product_id Not Like '%-NS'
      AND product_id Not Like '%-NS'
      AND description2 Not Like '% OF %'
      AND description3 Not Like '% OF %'
      AND description3 Not Like '%REPLACEMENT%'
      AND product_id Not Like '%-1'
      AND product_id Not Like '%-2'
      AND product_id Not Like '%-3'

    UNION

    SELECT *
    FROM {table_name}
    WHERE product_description Like '%SET%'
      OR product_description Like '%SINGLE%'
    '''

#Return SQL Select statement CASE,WHEN,END AS to create new fields 'In stock?' and 'Regular price'
def sql_get_stock_price(table_name):
    return f'''
    SELECT {table_name}.product_id, 
      CASE 
        WHEN {table_name}.avail > 0 AND ({table_name}.reg_price > 2 OR {table_name}.pu_price > 2) THEN 1
        ELSE 0
      END AS `In stock?`,
      CASE 
        WHEN {table_name}.reg_price > {table_name}.pu_price THEN {table_name}.reg_price
        ELSE {table_name}.pu_price
      END AS `Regular price`
    FROM {table_name} 
    '''

#Return SQL Select statement with conditions to find discontinuing items on the website that are out of stock
def sql_select_items_to_remove(website_items, valid_items, valid_items_stock_avail):
    return f'''
      SELECT {website_items}.id, {website_items}.name, {valid_items_stock_avail}.`In Stock?`, {valid_items_stock_avail}.`Regular price`
      FROM website_items, valid_items, valid_items_stock_avail
        WHERE {website_items}.name = {valid_items}.product_id
          AND {website_items}.name = {valid_items_stock_avail}.product_id
          AND {valid_items_stock_avail}.`In Stock?`= 0
          AND {valid_items}.is_continuing = 0
    '''

#Return SQL Select statement with conditions to find items that need to be updated due to out-dated price or inStock status
def sql_select_items_to_update(website_items, valid_items_stock_avail):
    return f'''
    SELECT {website_items}.id, {website_items}.name, {valid_items_stock_avail}.`In Stock?`, {valid_items_stock_avail}.`Regular price`
    FROM website_items, valid_items_stock_avail
    WHERE {website_items}.name = {valid_items_stock_avail}.product_id
      AND ({website_items}.in_stock != {valid_items_stock_avail}.`In Stock?`
        OR {website_items}.reg_price != {valid_items_stock_avail}.`Regular price`)
    '''


#===================  DEBUGGING FUNCTIONS ========================================

#Creates '`query`.csv' file from passed in query as a list
def create_csv_query(cursor, query_function):
      
    query_dict = {
      'sql_select_items_to_remove' : sql_select_items_to_remove('website_items', 'valid_items', 'valid_items_stock_avail'),
      'sql_select_items_to_update' : sql_select_items_to_update('website_items', 'valid_items_stock_avail')
    }

    print(f'Creating Query: `{query_function}`.csv')

    with open(f"G:/Kevin/website info/Zach's Toolbox/Query_{query_function}.csv", "w", newline='') as csvfile:
      spamwriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
      cursor.execute('{};'.format(query_dict.get(query_function)))
      for row in cursor.fetchall():
        spamwriter.writerow([cell for cell in row])

#Creates '`table_name`.csv' file from passed in table_name
def create_csv_table(cursor, table_name):
    with open(f"G:\\Kevin\\website info\\Zach's Toolbox\\{table_name}.csv", "w", newline='') as csvfile:
      spamwriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
      cursor.execute(f'SELECT * FROM {table_name};')
      for row in cursor.fetchall():
        spamwriter.writerow([cell for cell in row])

#=================================================================================

#Gets Names/IDs and inserts into Table: 'website_items'
def create_table_woocommerce_items(cursor, wcapi, per_page = 100):
      
    #Create new 'website_items' table
    try:
      print(f"Creating `website_items` Table")
      create_table(cursor, 'website_items')
    except mysql.connector.Error as error:  print_error(error)

    total_items = 0

    response = wcapi.get('products')
    total_pages = int((int(response.headers['X-WP-TotalPages']) * 10) / per_page) + 1
    for i in range(1, total_pages + 1):
      print("{}%...".format(int((i * 100)/total_pages)))
      responseGET = wcapi.get('products?per_page={}&page='.format(per_page) + str(i)).json()
      for name in responseGET:
        total_items += 1
        insert_table(cursor, 'website_items', (name.get("id"), name.get("name"), int(name.get("stock_status") == 'instock'), name.get("regular_price")))

    cursor.execute('SELECT * FROM website_items;')
    
    if total_items > cursor.rowcount: 
      print("Error: Missing items due to fetched Duplicates from WooREST API")
      print(f"Missing {total_items - cursor.rowcount} items")

      #Automatically attempt to recover missing items once
      if per_page > 98:
        print("Attempting to Recover Missing Items...")
        create_table_woocommerce_items(cursor, wcapi, per_page - 2)

      #Provide User the option to proceed without missing item if already attempted to recover
      # else:
          

    else: print(f"Total Website Items: {total_items}")
  
    conn.commit()
  
#Get Accuterm's Product Info from Exported Excel Workbook
def create_tables_accuterm_items(cursor, file):

  #Dictionary of Column Field Names : Column Index
  col_dict = {
    'Product Description' : 4,
    'Description #2' : 5,
    'Description #3' : 6,
    'Reg Price' : 7,
    'P/U Price' : 8,
    'Avail' : 11,
  }

  #raw_items table serves as a placeholder for extracted products in excel sheet(s)
  create_table(cursor, 'raw_items')

  #Load worksheet object from file path
  wb_obj = load_workbook(filename=file)
  for sheet in wb_obj.sheetnames:
    wsheet = wb_obj[sheet]

    #Copy products from each worksheet into raw_items table
    for key, *values in wsheet.iter_rows():
      if key.value != ' ':
            
        data_row = [v.value for v in values if v.column in col_dict.values() and v.row > 2]
        data_row = ['' if data == None else data for data in data_row]
        data_row.insert(0, key.value)

        #Insert True in field:'is_continuing' if item is continuing and false otherwise
        if wb_obj.sheetnames.index(sheet) <= 1: data_row.append(True)
        else:                                   data_row.append(False)

        insert_table(cursor, 'raw_items', tuple(data_row))

  #Copy/filter raw_items into valid_items table once reaches last sheet
  if wb_obj.sheetnames.index(sheet) == (len(wb_obj.sheetnames) - 1): 
    
    cursor.execute('CREATE TABLE valid_items {};'.format(sql_get_valid_items('raw_items')))

    #Truncates raw_item table
    cursor.execute('TRUNCATE TABLE raw_items;')

    #Insert into valid_items_stock_avail table with each items inStock status and price from valid_items table
    cursor.execute('CREATE TABLE valid_items_stock_avail {};'.format(sql_get_stock_price('valid_items')))

  conn.commit()

def delete_items_to_remove(cursor, wcapi):

    batch_data = {"delete" : []}
    
    #Select ID's of discontinued items that are no longer in stock or have prices < $2 
    cursor.execute('{};'.format(sql_select_items_to_remove('website_items', 'valid_items', 'valid_items_stock_avail')))
    for row in cursor.fetchall():
        batch_data["delete"].append(row[0])
    
    #Delete batch of selected ID's from website
    response = wcapi.post('product', batch_data)

#=================================================================================

# product_data = {
#     "name": "Zach's testing Product 2",
#     "type": "simple",
#     "regular_price": "1234",
#     "In stock?": 0,
#     "short_description": "testing DO NOT PURCHASE",
#     "description": "<p>Bring your living space to entirely new and fashionable heights, with this counter-height dining bench. This piece is button tufted and accented with its sturdy wooden construction will enable years of enjoyment. Its soft, solid linen upholstery tops it off like icing on a cake. Available in <em>Grey and Beige</em> colors to choose.</p><p><strong>Features:</strong></p><ul><li>Made of wood, MDF, and linen fabric</li><li>Includes 1 bench only</li><li>Solid medium linen fabric upholstery</li><li>Button-tufted, cushioned top for cozy comfort</li><li>Weathered Grey Finish on wood frame</li><li>Trestle Based</li></ul><p><strong>Overall Dimensions: (L x W x H)</strong></p><ul><li>Dimensions: 48 inches wide x 16 inches deep x 19 inches tall</li><li>Weight capacity: 250 pounds</li></ul>",
#     "categories": [
#       {
#         "id": 654
#       }
#     ],

#     "images": [
#         {
#             "src": "https://bestqualityfurn.com/wp-content/uploads/VEN-V-1-scaled.jpg",
#             "alt": "Under ZackaDaHacka's test"
#         }
#     ]
# }
# response = wcapi.post('products', product_data)
#CreateCategoryCSV(wcapi)

try:
  print ("My SQL Connection established.")
  cursor = conn.cursor(buffered=True)

  # drop_tables(cursor)

  # create_table_woocommerce_items(cursor, wcapi)

  # wcapi.delete("products/5768")

  # create_tables_accuterm_items(cursor, 'G:/Kevin/website info/New Accuterm Info/082322 inventory.xlsx')

  create_csv_table(cursor, 'website_items')
  create_csv_table(cursor, 'valid_items')
  create_csv_table(cursor, 'valid_items_stock_avail')


  create_csv_query(cursor, 'sql_select_items_to_remove')
  create_csv_query(cursor, 'sql_select_items_to_update')

except mysql.connector.Error as err:
  print("Something went wrong: {}".format(err))
