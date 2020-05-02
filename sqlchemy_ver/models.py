from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column,Integer,String
from sqlalchemy.orm import sessionmaker


engine = create_engine('sqlite:///sql.db')
Base = declarative_base()

class User(Base):
    __tablename__='users'
    id = Column(Integer,primary_key=True,autoincrement=True)
    name = Column(String)
    fullname = Column(String)
    
db_session = sessionmaker(bind=engine)
session = db_session()

Base.metadata.create_all(engine)

user1 = User(name='liyang')

session.add(user1)

session.commit()
