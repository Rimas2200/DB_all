import mysql.connector   #pip install mysql-connector-python

mysql_connection = mysql.connector.connect(
    host='127.0.0.1',
    user='root',
    password='',
    database='timetable'
)
cursor = mysql_connection.cursor()

create_table_query = '''
        CREATE TABLE IF NOT EXISTS group_name(
        id INT AUTO_INCREMENT PRIMARY KEY,
        group_name TEXT,
        direction_abbreviation TEXT
    );
    CREATE TABLE IF NOT EXISTS professor(
        id INT AUTO_INCREMENT PRIMARY KEY,
        last_name TEXT,
        first_name TEXT,
        middle_name TEXT,
        position TEXT,
        departament TEXT
    );
    CREATE TABLE IF NOT EXISTS direction(
        id INT AUTO_INCREMENT PRIMARY KEY,
        direction_abbreviation TEXT,
        name TEXT,
        faculty TEXT
    );
    CREATE TABLE IF NOT EXISTS discipline(
        id INT AUTO_INCREMENT PRIMARY KEY,
        discipline_name TEXT
    );
    CREATE TABLE IF NOT EXISTS classroom(
        id INT AUTO_INCREMENT PRIMARY KEY,
        room_number TEXT,
        building TEXT
    );
    CREATE TABLE IF NOT EXISTS couple_type(
        id INT AUTO_INCREMENT PRIMARY KEY,
        pair_type TEXT,
        name TEXT,
        faculty TEXT
    );
    CREATE TABLE IF NOT EXISTS address(
        id INT AUTO_INCREMENT PRIMARY KEY,
        address TEXT,
        faculty TEXT
    );
    CREATE TABLE IF NOT EXISTS departament(
        id INT AUTO_INCREMENT PRIMARY KEY,
        name TEXT,
        phone TEXT,
        faculty TEXT
    );
'''
cursor.execute(create_table_query)

cursor.close()
mysql_connection.close()

mysql_connection = mysql.connector.connect(
    host='localhost',
    user='root',
    password='',
    database='timetable'
)
cursor = mysql_connection.cursor()

cursor.close()
mysql_connection.close()