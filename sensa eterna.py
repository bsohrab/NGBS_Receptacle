
from configparser import ConfigParser
import psycopg2


def config(filename='dbinfo.ini', section='postgresql'):
    #the parser for the info from the config file
    parser = ConfigParser()
    parser.read(filename)

    db= {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in hte {1} file'.format(section,filename))
    return db

def connect():
    """Connect to the PostgresSQL database server"""

    conn = None
    try:
        params = config()
        print("Attempting the connection")
        conn = psycopg2.connect(**params)

        cur = conn.cursor()

        print('PostgreSQL Database Version:')
        cur.execute('SELECT version()')

        db_version = cur.fetchone()
        print(db_version)

        create_tables(cur)

        cur.close()
        conn.commit()
    except(Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Datanase connection closed all fully and stuff')

def create_tables(cur):
    """Create tables in the PostgreSQL"""
    commands = (
        """
        CREATE TABLE vendors (
        vendor_id SERIAL PRIMARY KEY
        vendor_name VARCHAR(255) NOT NULL
        )
        """,
         """
        CREATE TABLE part_drawings (
                part_id INTEGER PRIMARY KEY,
                file_extension VARCHAR(5) NOT NULL,
                drawing_data BYTEA NOT NULL,
                FOREIGN KEY (part_id)
                REFERENCES parts (part_id)
                ON UPDATE CASCADE ON DELETE CASCADE
        )
        """,
        """CREATE TABLE NGBS_DATA(
            project_id VARCHAR(15) PRIMARY KEY,
            zip_code INTEGER NOT NULL,
            certification_level VARCHAR(15)
            )
        """

    )
    for command in commands:
        cur.execute(command)


def gather_data(cur):
    """Add the data from the device to the psql lord database"""
    command=(
        """
        COPY NGBSDATAEXTRACT FROM "C:/Users/spasikhani/Documents/mass production test/45.csv" WITH CSV HEADER"""
    )
if __name__ == '__main__':
    connect()
