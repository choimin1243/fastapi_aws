from sqlalchemy import Boolean, Column, Integer, String, ForeignKey
from sqlalchemy.orm import relationship
from database import Base

class Todos(Base):
    __tablename__ = "package"

    id = Column(Integer, primary_key=True, index=True)
    title = Column(String)
