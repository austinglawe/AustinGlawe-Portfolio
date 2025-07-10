# -----------------------------------------
# Python_Databases:
# Working with databases
# -----------------------------------------
#
# SQLite example:
#   import sqlite3
#   conn = sqlite3.connect('example.db')
#   cursor = conn.cursor()
#   cursor.execute('CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, name TEXT)')
#   cursor.execute('INSERT INTO users (name) VALUES (?)', ('Austin',))
#   conn.commit()
#   cursor.execute('SELECT * FROM users')
#   rows = cursor.fetchall()
#   for row in rows:
#       print(row)
#   conn.close()
#
# pyodbc example:
#   import pyodbc
#   conn = pyodbc.connect('DRIVER={SQL Server};SERVER=server_name;DATABASE=db_name;UID=user;PWD=password')
#   cursor = conn.cursor()
#   cursor.execute('SELECT * FROM table_name')
#   for row in cursor:
#       print(row)
#   conn.close()
#
# SQLAlchemy ORM example:
#   from sqlalchemy import create_engine, Column, Integer, String
#   from sqlalchemy.ext.declarative import declarative_base
#   from sqlalchemy.orm import sessionmaker
#   engine = create_engine('sqlite:///example.db')
#   Base = declarative_base()
#
#   class User(Base):
#       __tablename__ = 'users'
#       id = Column(Integer, primary_key=True)
#       name = Column(String)
#
#   Base.metadata.create_all(engine)
#   Session = sessionmaker(bind=engine)
#   session = Session()
#   new_user = User(name="Austin")
#   session.add(new_user)
#   session.commit()
#   for user in session.query(User).all():
#       print(user.name)
#
# Best practices:
# - Use parameterized queries.
# - Close connections properly.
# - Use ORM for complex projects.
# - Handle DB exceptions.
#
# -----------------------------------------
