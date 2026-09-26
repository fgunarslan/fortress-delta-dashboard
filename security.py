import os,base64,hashlib,hmac,secrets
from datetime import datetime,timezone
from db import SessionLocal,User,audit

APP_SECRET=os.getenv("APP_SECRET","")
ADMIN_USER="fuat"
ADMIN_PASSWORD=os.getenv("ADMIN_PASSWORD","")

def now():
    return datetime.now(timezone.utc)

def hash_password(password):
    salt=secrets.token_bytes(16)
    rounds=240000
    dk=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,rounds)
    return f"pbkdf2_sha256${rounds}$"+base64.urlsafe_b64encode(salt).decode()+"$"+base64.urlsafe_b64encode(dk).decode()

def verify_password(password,encoded):
    try:
        _,rounds,salt64,dk64=encoded.split("$",3)
        salt=base64.urlsafe_b64decode(salt64.encode())
        expected=base64.urlsafe_b64decode(dk64.encode())
        got=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,int(rounds))
        return hmac.compare_digest(got,expected)
    except Exception:
        return False

def sign_session(username):
    if not APP_SECRET:
        raise RuntimeError("APP_SECRET is not configured")
    payload=f"{username}|{int(now().timestamp())}"
    sig=hmac.new(APP_SECRET.encode(),payload.encode(),hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}|{sig}".encode()).decode()

def read_session(token):
    try:
        raw=base64.urlsafe_b64decode(token.encode()).decode()
        username,ts,sig=raw.rsplit("|",2)
        payload=f"{username}|{ts}"
        expected=hmac.new(APP_SECRET.encode(),payload.encode(),hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig,expected): return None
        if int(now().timestamp())-int(ts)>60*60*24*7: return None
        return username
    except Exception:
        return None

def ensure_admin():
    if not ADMIN_PASSWORD:
        raise RuntimeError("ADMIN_PASSWORD must be configured before first start.")
    with SessionLocal() as db:
        # Only Fuat may ever be admin.
        for u in db.query(User).filter(User.username!=ADMIN_USER,User.role=="admin").all():
            u.role="viewer"
        fuat=db.query(User).filter_by(username=ADMIN_USER).first()
        if fuat is None:
            fuat=User(username=ADMIN_USER,password_hash=hash_password(ADMIN_PASSWORD),role="admin",active=True)
            db.add(fuat)
        else:
            fuat.role="admin"; fuat.active=True
        db.commit()
