import pandas as pd
import cx_Oracle


# Function to insert data into Oracle
def insert_data_to_oracle(dataframe, connection, Test1):
    cursor = connection.cursor()
    # Generate placeholders dynamically based on the number of columns
    placeholders = ", ".join([":" + str(i + 1) for i in range(len(dataframe.columns))])
    insert_query = f"INSERT INTO {Test1} ({', '.join(dataframe.columns)}) VALUES ({placeholders})"
    
    # Iterate through DataFrame rows and insert each row into the table
    for row in dataframe.itertuples(index=False, name=None):
        cursor.execute(insert_query, row)
    
    connection.commit()
    print("Data inserted successfully.")
    cursor.close()


# Step 1: Read data from Excel
excel_file = "input_file.xlsx"  # Replace with your file path
df = pd.read_excel(excel_file)

# Add the path to the Oracle Instant Client
cx_Oracle.init_oracle_client(lib_dir=r"#add your path\instantclient_19_8")

# Database connection details
user = "ABC"
password = "12345"
dsn = "db_ip_address"

# Connect to the Oracle database
connection = cx_Oracle.connect(user=user, password=password, dsn=dsn)
print(connection)

# Step 3: Call the function to insert data
table_name = "Test1"  # Replace with your target table name
insert_data_to_oracle(df, connection, table_name)

# Step 4: Close the database connection
connection.close()
