import os
import psycopg2
from psycopg2.pool import SimpleConnectionPool
from dotenv import load_dotenv
load_dotenv()  

pool = None
def init_db_pool():
    global pool
    if pool is None:
        pool = SimpleConnectionPool(
            15,                
            20,               
            host=os.environ.get("DB_HOST"),
            database=os.environ.get("DB_NAME"),
            user=os.environ.get("DB_USER"),
            password=os.environ.get("DB_PASS"),
            port=os.environ.get("DB_PORT", 5432),
            connect_timeout=5
        )
    return pool


# def get_db_conn():
#     pool = init_db_pool()
#     return pool.getconn()   

# def release_db_conn(conn):
#     pool = init_db_pool()
#     pool.putconn(conn)

def get_db_conn():
    pool = init_db_pool()
    conn = pool.getconn()
    conn.autocommit = True
    return conn

def release_db_conn(conn):
    try:
        if not conn.closed:
            conn.rollback()
    finally:
        pool = init_db_pool()
        pool.putconn(conn)
