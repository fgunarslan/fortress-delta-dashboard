import os, json
from datetime import datetime, timezone
from sqlalchemy import create_engine, Column, Integer, String, Float, Boolean, DateTime, Text
from sqlalchemy.orm import declarative_base, sessionmaker

DATABASE_URL=os.getenv("DATABASE_URL","sqlite:///./fortress_delta_v2.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL="postgresql+psycopg://"+DATABASE_URL[len("postgres://"):]
elif DATABASE_URL.startswith("postgresql://") and "+psycopg" not in DATABASE_URL:
    DATABASE_URL="postgresql+psycopg://"+DATABASE_URL[len("postgresql://"):]

connect_args={"check_same_thread":False} if DATABASE_URL.startswith("sqlite") else {}
engine=create_engine(DATABASE_URL,pool_pre_ping=True,connect_args=connect_args)
SessionLocal=sessionmaker(bind=engine,autoflush=False,autocommit=False)
Base=declarative_base()

def utcnow():
    return datetime.now(timezone.utc)

class User(Base):
    __tablename__="users"
    id=Column(Integer,primary_key=True)
    username=Column(String(80),unique=True,index=True,nullable=False)
    password_hash=Column(String(255),nullable=False)
    role=Column(String(20),nullable=False,default="viewer")
    active=Column(Boolean,default=True,nullable=False)
    created_at=Column(DateTime(timezone=True),default=utcnow)
    updated_at=Column(DateTime(timezone=True),default=utcnow)

class Setting(Base):
    __tablename__="settings"
    key=Column(String(80),primary_key=True)
    value=Column(Text,nullable=False)

class Position(Base):
    __tablename__="positions"
    id=Column(Integer,primary_key=True)
    ticker=Column(String(30),nullable=False)
    expiry=Column(String(10),nullable=False)
    option_type=Column(String(1),nullable=False)
    strike=Column(Float,nullable=False)
    quantity=Column(Float,nullable=False)
    multiplier=Column(Float,nullable=False,default=100)
    bloomberg_security=Column(String(180),nullable=False)
    underlying_security=Column(String(120),nullable=False)
    active=Column(Boolean,default=True,nullable=False)
    source=Column(String(30),default="manual")
    created_at=Column(DateTime(timezone=True),default=utcnow)
    updated_at=Column(DateTime(timezone=True),default=utcnow)

class Audit(Base):
    __tablename__="audit"
    id=Column(Integer,primary_key=True)
    ts=Column(DateTime(timezone=True),default=utcnow,index=True)
    actor=Column(String(80),nullable=False)
    action=Column(String(80),nullable=False)
    detail=Column(Text,nullable=False)

class Snapshot(Base):
    __tablename__="snapshots"
    id=Column(Integer,primary_key=True)
    ts=Column(DateTime(timezone=True),default=utcnow,index=True)
    payload=Column(Text,nullable=False)

class ImportJob(Base):
    __tablename__="import_jobs"
    id=Column(String(64),primary_key=True)
    created_at=Column(DateTime(timezone=True),default=utcnow)
    filename=Column(String(255),nullable=False)
    headers_json=Column(Text,nullable=False)
    rows_json=Column(Text,nullable=False)

Base.metadata.create_all(engine)

def get_setting(db,key,default=None):
    x=db.get(Setting,key)
    return x.value if x else default

def set_setting(db,key,value):
    x=db.get(Setting,key)
    if x is None:
        db.add(Setting(key=key,value=str(value)))
    else:
        x.value=str(value)

def audit(db,actor,action,detail):
    db.add(Audit(actor=actor,action=action,detail=detail))

def config_bump(db,actor,reason):
    v=int(get_setting(db,"config_version","1"))+1
    set_setting(db,"config_version",v)
    set_setting(db,"config_updated_at",utcnow().isoformat())
    audit(db,actor,"CONFIG_PUBLISHED",f"v{v}: {reason}")
    return v
