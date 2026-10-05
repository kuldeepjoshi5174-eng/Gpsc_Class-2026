"""GPSC Portal - FastAPI backend.

Roles
  admin    - exactly one account; creates/edits staff, can create, edit and delete everything
  faculty  - adds lecture attendance (L1/L2) and test marks for a date; cannot edit saved entries
  library  - adds library attendance (in/out/book) for a date; cannot edit saved entries
Run:  uvicorn main:app --host 0.0.0.0 --port 8000
"""
import base64, hashlib, hmac, json, os, re, secrets, sqlite3, time
from collections import Counter
from contextlib import contextmanager
from datetime import date as _date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import List, Literal, Optional

from fastapi import Depends, FastAPI, Header, HTTPException, Query
from fastapi.responses import FileResponse, Response
from pydantic import BaseModel, Field, field_validator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB_PATH = os.environ.get("DB_PATH", str(BASE / "portal.db"))
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
USE_PG = DATABASE_URL.startswith(("postgres://", "postgresql://"))
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-change-me").encode()
ADMIN_USERNAME = (os.environ.get("ADMIN_USERNAME") or "admin").strip().lower()
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD") or os.environ.get("TEACHER_PASSWORD") or ""
RESET_ADMIN = os.environ.get("RESET_ADMIN_PASSWORD") == "1"   # emergency: reset admin password from env on start
TOKEN_TTL = 12 * 3600
IST = timezone(timedelta(hours=5, minutes=30))

def today_ist() -> str:
    return datetime.now(IST).date().isoformat()

def now_ist() -> str:
    return datetime.now(IST).isoformat(timespec="seconds")

if USE_PG:
    import psycopg
    from psycopg.rows import dict_row
INTEGRITY = (sqlite3.IntegrityError,) + ((psycopg.errors.IntegrityError,) if USE_PG else ())

STUDENTS = [
 (1,"Janki Rameshbhai Lodha (લોઢા જાનકી)","G.D. Modi College"),(2,"Priyanka Amrutbhai Desai (દેસાઈ પ્રિયંકા)","G.D. Modi College"),
 (3,"Krish Ashokkumar Badani (બદાણી ક્રિશ)","G.D. Modi College"),(4,"Sheetal Tanaji Thakor (ઠાકોર શિતલ)","G.D. Modi College"),
 (5,"Dhruvi Kishorbhai Rathod (રાઠોડ ધ્રુવી)","M.A. Parikh College"),(6,"Piyush Prakashbhai Gohil (ગોહિલ પીયુષ)","G.D. Modi College"),
 (7,"Parth Bharatkumar Trivedi (ત્રિવેદી પાર્થ)","M.A. Parikh College"),(8,"Tushar Pareshbhai Bhatiya (ભાટિયા તુષાર)","M.A. Parikh College"),
 (9,"Mohammad Imran Pathan (પઠાણ મોહમ્મદ)","M.A. Parikh College"),(10,"Prit Ashokkumar Chorasiya (ચોરાસિયા પ્રિત)","M.A. Parikh College"),
 (11,"Shubham Amitkumar Bhatiya (ભાટિયા શુભમ)","M.A. Parikh College"),(12,"Namita Sardarbhai Chaudhary (ચૌધરી નમિતા)","M.A. Parikh College"),
 (13,"Amrit Dineshji Mali (માળી અમ્રીત)","M.A. Parikh College"),(14,"Khushal Dhanaji Prajapati (પ્રજાપતિ ખુશાલ)","M.A. Parikh College"),
 (15,"Agres Karmibhai Luni (લુણી અગ્રેસ)","M.A. Parikh College"),(16,"Sarin Yashvantbhai Chaudhary (ચૌધરી સરીન)","M.A. Parikh College"),
 (17,"Nilay Nareshkumar Patel (પટેલ નિલય)","Smt B.K. Mehta BCA"),(18,"Akki Kiranbhai Solanki (સોલંકી અક્કી)","R.R. Mehta College"),
 (19,"Hetal Amratji Koli (કોળી હેતલ)","M.A. Parikh College"),(20,"Kaynat Sipahi (સિપાહી કાયનાત)","HNGU Patan"),
 (21,"Darshan Sevantilal Darji (દરજી દર્શન)","HNGU Patan"),(22,"Naresh Panchaji Rabari (રબારી નરેશ)","BAOU"),
 (23,"Saleha Jabirbhai Sipahi (સિપાહી સાલેહા)","HNGU Patan"),(24,"Sapna Nanjibhai Judal (જુડાલ સપના)","M.A. Parikh College"),
 (25,"Usha Kamaji Prajapati (પ્રજાપતિ ઉષા)","M.A. Parikh College"),(26,"Ansuya Vagjibhai Makwana (મકવાણા અનસુયા)","M.A. Parikh College"),
 (27,"Shilpa Kanjibhai Thakor (ઠાકોર શિલ્પા)","M.A. Parikh College"),(28,"Sanjana Bachubhai Raval (રાવળ સંજના)","C.L. Parikh Commerce"),
 (29,"Ashish Narayanbhai Patni (પટની આશિષ)","M.A. Parikh College"),(30,"Abrar Kasambhai Umatiya (ઉમાતિયા અબરાર)","G.D. Modi College"),
 (31,"Ovez Adambhai Umatiya (ઉમતીયા ઓવેઝ)","G.D. Modi College"),(32,"Urvashi Dharmendrabhai Kuvariya (કુંવરીયા ઉર્વશી)","G.D. Modi College"),
 (33,"Parikshit Daljibhai Chaudhary (ચૌધરી પરીક્ષિત)","G.D. Modi College"),(34,"Sonal Dashrathbhai Gelot (ગેલોત સોનલ)","G.D. Modi College"),
 (35,"Sakshi Rameshbhai Goswami (ગોસ્વામી સાક્ષી)","G.D. Modi College"),(36,"Himanshi Rameshji Thakor (ઠાકોર હિમાંશી)","G.D. Modi College"),
 (37,"Karan Mukeshji Thakor (ઠાકોર કરણ)","G.D. Modi College"),(38,"Parth Dhirajbhai Parmar (પરમાર પાર્થ)","R.R. Mehta College"),
 (39,"Farhan Iqbalbhai Sindhi (સિંધી ફરહાન)","G.D. Modi College"),(40,"Muskan Iqbalbhai Sindhi (સિંધી મુસ્કાન)","G.D. Modi College"),
 (41,"Shubham Sureshbhai Chauhan (ચૌહાણ શુભમ)","R.R. Mehta College"),(42,"Neha Prakashbhai Prajapati (પ્રજાપતિ નેહા)","M.A. Parikh College"),
 (43,"Rahul Chetankumar Lakhavara (લખવારા રાહુલ)","G.D. Modi College"),(44,"Sankalp Hareshbhai Shrimali (શ્રીમાળી સંકલ્પ)","M.A. Parikh College"),
 (45,"Alviz Iqbalbhai Jagirdar (જાગીરદાર એલવીજ)","B.L. Parikh BBA"),(46,"Alnaz Iqbalbhai Jagirdar (જાગીરદાર એલનાજ)","B.L. Parikh BBA"),
 (47,"Sourav Jitendrakumar Joshi (જોષી સૌરવ)","R.R. Mehta College"),(48,"Mahir Zakirbhai Kureshi (કુરેશી માહીર)","C.L. Parikh Commerce"),
 (49,"Anil Bhikhabhai Parmar (પરમાર અનિલ)","R.R. Mehta College"),(50,"Dipika Rameshbhai Makwana (મકવાણા દિપિકા)","G.D. Modi College"),
 (51,"Mital Pareshbhai Prajapati (પ્રજાપતિ મિતલ)","R.R. Mehta College"),
]

_ID = "SERIAL PRIMARY KEY" if USE_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
SCHEMA = f"""
CREATE TABLE IF NOT EXISTS staff(
  id {_ID}, username TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
  role TEXT NOT NULL CHECK(role IN ('admin','faculty','library')),
  subject TEXT, pw_hash TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1, created_at TEXT);
CREATE TABLE IF NOT EXISTS batches(
  id {_ID}, name TEXT NOT NULL UNIQUE, description TEXT, active INTEGER NOT NULL DEFAULT 1, created_at TEXT);
CREATE TABLE IF NOT EXISTS staff_batches(
  staff_id INTEGER NOT NULL REFERENCES staff(id), batch_id INTEGER NOT NULL REFERENCES batches(id),
  PRIMARY KEY(staff_id, batch_id));
CREATE TABLE IF NOT EXISTS students(
  id {_ID}, batch_id INTEGER NOT NULL REFERENCES batches(id), roll INTEGER NOT NULL CHECK(roll > 0),
  name TEXT NOT NULL, college TEXT NOT NULL DEFAULT '', mobile TEXT, parent_name TEXT, parent_mobile TEXT,
  active INTEGER NOT NULL DEFAULT 1, created_at TEXT, UNIQUE(batch_id, roll));
CREATE TABLE IF NOT EXISTS sessions(
  batch_id INTEGER NOT NULL REFERENCES batches(id), date TEXT NOT NULL,
  slot TEXT NOT NULL CHECK(slot IN ('L1','L2','LIB')),
  faculty_id INTEGER REFERENCES staff(id), subject TEXT, topic TEXT, updated_at TEXT,
  PRIMARY KEY(batch_id, date, slot));
CREATE TABLE IF NOT EXISTS att(
  batch_id INTEGER NOT NULL, date TEXT NOT NULL, slot TEXT NOT NULL,
  student_id INTEGER NOT NULL REFERENCES students(id),
  status TEXT NOT NULL CHECK(status IN ('Present','Absent')),
  lib_in TEXT, lib_out TEXT, book TEXT,
  PRIMARY KEY(batch_id, date, slot, student_id),
  FOREIGN KEY(batch_id, date, slot) REFERENCES sessions(batch_id, date, slot) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS tests(
  id {_ID}, batch_id INTEGER NOT NULL REFERENCES batches(id), date TEXT NOT NULL, title TEXT NOT NULL,
  max_marks DOUBLE PRECISION NOT NULL CHECK(max_marks > 0), created_by INTEGER,
  UNIQUE(batch_id, date, title));
CREATE TABLE IF NOT EXISTS marks(
  test_id INTEGER NOT NULL REFERENCES tests(id) ON DELETE CASCADE,
  student_id INTEGER NOT NULL REFERENCES students(id), score DOUBLE PRECISION,
  PRIMARY KEY(test_id, student_id));
"""

class Conn:
    """Same SQL ('?' placeholders) runs on SQLite (local) and Postgres (cloud)."""
    def __init__(self, raw): self.raw = raw
    def _q(self, sql): return sql.replace("?", "%s") if USE_PG else sql
    def execute(self, sql, params=()): return self.raw.execute(self._q(sql), params)
    def executemany(self, sql, seq):
        if USE_PG:
            cur = self.raw.cursor(); cur.executemany(self._q(sql), seq); return cur
        return self.raw.executemany(sql, seq)

@contextmanager
def db():
    if USE_PG:
        raw = psycopg.connect(DATABASE_URL, row_factory=dict_row, prepare_threshold=None)
    else:
        raw = sqlite3.connect(DB_PATH, timeout=15)
        raw.row_factory = lambda c, r: {d[0]: v for d, v in zip(c.description, r)}
        raw.execute("PRAGMA foreign_keys=ON")
        raw.execute("PRAGMA journal_mode=WAL")
    con = Conn(raw)
    try:
        yield con
        raw.commit()
    except Exception:
        raw.rollback()
        raise
    finally:
        raw.close()

# ---------- passwords & tokens ----------
def hash_pw(pw: str) -> str:
    salt = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 200_000).hex()
    return f"pbkdf2${salt}${h}"

def check_pw(pw: str, stored: str) -> bool:
    try:
        _, salt, h = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 200_000).hex()
        return hmac.compare_digest(calc, h)
    except Exception:
        return False

def _b64(b: bytes) -> str: return base64.urlsafe_b64encode(b).decode().rstrip("=")

def make_token(uid: int) -> str:
    payload = _b64(json.dumps({"u": uid, "e": int(time.time()) + TOKEN_TTL}).encode())
    return f"{payload}.{hmac.new(SECRET_KEY, payload.encode(), hashlib.sha256).hexdigest()}"

def parse_token(authorization: Optional[str]) -> Optional[int]:
    try:
        payload, sig = authorization.removeprefix("Bearer ").split(".")
        if not hmac.compare_digest(sig, hmac.new(SECRET_KEY, payload.encode(), hashlib.sha256).hexdigest()):
            return None
        data = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
        return data["u"] if data["e"] > time.time() else None
    except Exception:
        return None

def current_user(authorization: Optional[str] = Header(None)):
    uid = parse_token(authorization)
    if uid is None:
        raise HTTPException(401, "લોગિન જરૂરી છે / સેશન પૂરું થયું")
    with db() as con:   # re-read every time, so role/disable changes apply immediately
        u = con.execute("SELECT id,username,name,role,subject,active FROM staff WHERE id=?", (uid,)).fetchone()
    if not u or not u["active"]:
        raise HTTPException(401, "આ ખાતું બંધ છે")
    return u

def need(*roles):
    def dep(u=Depends(current_user)):
        if u["role"] not in roles:
            raise HTTPException(403, "આ કામ માટે તમારી પરવાનગી નથી")
        return u
    return dep

def public_user(u) -> dict:
    return {k: u[k] for k in ("id", "username", "name", "role", "subject")}

def batch_access(con, u, batch_id: int):
    """Admin: any batch. Faculty/library: only active batches assigned to them."""
    b = con.execute("SELECT * FROM batches WHERE id=?", (batch_id,)).fetchone()
    if not b:
        raise HTTPException(404, "બેચ મળ્યો નથી")
    if u["role"] != "admin" and (not b["active"] or not con.execute(
            "SELECT 1 FROM staff_batches WHERE staff_id=? AND batch_id=?", (u["id"], batch_id)).fetchone()):
        raise HTTPException(403, "આ બેચની તમારી પાસે પરવાનગી નથી")
    return b

# ---------- init / migration ----------
def _columns(con, table):
    if USE_PG:
        rows = con.execute("SELECT column_name AS name FROM information_schema.columns "
                           "WHERE table_name=? AND table_schema=current_schema()", (table,)).fetchall()
    else:
        rows = con.execute(f"PRAGMA table_info({table})").fetchall()
    return [r["name"] for r in rows]

def init_db():
    with db() as con:
        cols = _columns(con, "students")
        if cols and "batch_id" not in cols:      # database from the older single-batch version: keep it, start fresh
            for t in ("marks", "att", "tests", "sessions", "students"):
                if _columns(con, t):
                    con.execute(f"ALTER TABLE {t} RENAME TO legacy_{t}")
        for stmt in SCHEMA.split(";"):
            if stmt.strip():
                con.execute(stmt)
        admin = con.execute("SELECT id FROM staff WHERE role='admin'").fetchone()
        if not admin:
            if not ADMIN_PASSWORD:
                raise RuntimeError("ADMIN_PASSWORD (અથવા TEACHER_PASSWORD) સેટ કરો - પ્રથમ એડમિન બનાવવા માટે જરૂરી છે.")
            con.execute("INSERT INTO staff(username,name,role,subject,pw_hash,active,created_at) VALUES(?,?,?,?,?,1,?)",
                        (ADMIN_USERNAME, "એડમિન", "admin", "", hash_pw(ADMIN_PASSWORD), now_ist()))
        elif RESET_ADMIN and ADMIN_PASSWORD:
            con.execute("UPDATE staff SET pw_hash=? WHERE id=?", (hash_pw(ADMIN_PASSWORD), admin["id"]))
        if not con.execute("SELECT 1 FROM batches").fetchone():     # first start: create the default batch with the 51 students
            now = now_ist()
            con.execute("INSERT INTO batches(name,description,active,created_at) VALUES(?,?,1,?)", ("GPSC બેચ", "", now))
            bid = con.execute("SELECT id FROM batches WHERE name=?", ("GPSC બેચ",)).fetchone()["id"]
            con.executemany("INSERT INTO students(batch_id,roll,name,college,active,created_at) VALUES(?,?,?,?,1,?)",
                            [(bid, r, n, c, now) for r, n, c in STUDENTS])
            con.execute(f"INSERT INTO staff_batches(staff_id,batch_id) SELECT id, {int(bid)} FROM staff WHERE role<>'admin'")

app = FastAPI(title="GPSC Portal")
init_db()

# ---------- validation helpers / models ----------
TIME_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")
PHONE_RE = re.compile(r"^[0-9+\- ()]{0,20}$")
MONTH_RE = r"^\d{4}-(0[1-9]|1[0-2])$"
Slot = Literal["L1", "L2", "LIB"]
SLOT_ROLES = {"L1": ("faculty", "admin"), "L2": ("faculty", "admin"), "LIB": ("library", "admin")}

def _chk_date(v: str) -> str:
    try:
        _date.fromisoformat(v)
    except ValueError:
        raise ValueError("તારીખ ખોટી છે")
    return v

def _chk_phone(v):
    v = (v or "").strip()
    if not PHONE_RE.match(v):
        raise ValueError("મોબાઇલ નંબર ખોટો છે")
    return v

class Login(BaseModel):
    username: str = Field(max_length=60)
    password: str = Field(max_length=100)

class PwChange(BaseModel):
    old_password: str = Field(max_length=100)
    new_password: str = Field(min_length=6, max_length=100)

class UserIn(BaseModel):
    username: str = Field(min_length=3, max_length=30)
    name: str = Field(min_length=1, max_length=80)
    role: Literal["faculty", "library"]
    subject: str = Field("", max_length=80)
    password: str = Field(min_length=6, max_length=100)
    batch_ids: List[int] = []

    @field_validator("username")
    @classmethod
    def _u(cls, v):
        v = v.strip().lower()
        if not re.fullmatch(r"[a-z0-9._-]+", v):
            raise ValueError("username માં ફક્ત a-z, 0-9, . _ - વાપરો")
        return v

class UserUpd(BaseModel):
    name: Optional[str] = Field(None, min_length=1, max_length=80)
    role: Optional[Literal["faculty", "library"]] = None
    subject: Optional[str] = Field(None, max_length=80)
    active: Optional[bool] = None
    password: Optional[str] = Field(None, min_length=6, max_length=100)
    batch_ids: Optional[List[int]] = None

class BatchIn(BaseModel):
    name: str = Field(min_length=2, max_length=80)
    description: str = Field("", max_length=200)

class BatchUpd(BaseModel):
    name: Optional[str] = Field(None, min_length=2, max_length=80)
    description: Optional[str] = Field(None, max_length=200)
    active: Optional[bool] = None

class StudentIn(BaseModel):
    batch_id: int
    roll: Optional[int] = Field(None, ge=1, le=100000)
    name: str = Field(min_length=1, max_length=100)
    college: str = Field("", max_length=100)
    mobile: str = ""
    parent_name: str = Field("", max_length=100)
    parent_mobile: str = ""

    @field_validator("mobile", "parent_mobile")
    @classmethod
    def _m(cls, v): return _chk_phone(v)

class StudentUpd(BaseModel):
    roll: Optional[int] = Field(None, ge=1, le=100000)
    name: Optional[str] = Field(None, min_length=1, max_length=100)
    college: Optional[str] = Field(None, max_length=100)
    mobile: Optional[str] = None
    parent_name: Optional[str] = Field(None, max_length=100)
    parent_mobile: Optional[str] = None
    active: Optional[bool] = None

class BulkRow(BaseModel):
    roll: Optional[int] = Field(None, ge=1, le=100000)
    name: str = Field(max_length=100)
    college: str = Field("", max_length=100)
    mobile: str = ""

class BulkIn(BaseModel):
    batch_id: int
    rows: List[BulkRow] = Field(max_length=500)

class AttRow(BaseModel):
    student_id: int
    status: Literal["Present", "Absent"]
    lib_in: Optional[str] = None
    lib_out: Optional[str] = None
    book: Optional[str] = Field(None, max_length=120)

    @field_validator("lib_in", "lib_out")
    @classmethod
    def _t(cls, v):
        if v in (None, ""):
            return None
        if not TIME_RE.match(v):
            raise ValueError("સમય HH:MM ફોર્મેટમાં હોવો જોઈએ")
        return v

class SessionIn(BaseModel):
    batch_id: int
    date: str
    slot: Slot
    subject: str = Field("", max_length=80)
    topic: str = Field("", max_length=120)
    faculty_id: Optional[int] = None      # admin only: who took this session
    rows: List[AttRow]

    @field_validator("date")
    @classmethod
    def _d(cls, v): return _chk_date(v)

class MarkRow(BaseModel):
    student_id: int
    score: Optional[float] = None         # None = Absent

class TestIn(BaseModel):
    test_id: Optional[int] = None         # set = editing an existing test (admin only)
    batch_id: int
    date: str
    title: str = Field(min_length=1, max_length=120)
    max_marks: float = Field(gt=0, le=1000)
    faculty_id: Optional[int] = None      # admin only
    marks: List[MarkRow]

    @field_validator("date")
    @classmethod
    def _d(cls, v): return _chk_date(v)

# ---------- login / account ----------
_fails = {}
@app.post("/api/login")
def login(body: Login):
    key, now = body.username.strip().lower(), time.time()
    recent = [t for t in _fails.get(key, []) if now - t < 300]
    if len(recent) >= 8:
        raise HTTPException(429, "ઘણા ખોટા પ્રયત્નો. 5 મિનિટ પછી ફરી પ્રયત્ન કરો.")
    with db() as con:
        u = con.execute("SELECT * FROM staff WHERE username=?", (key,)).fetchone()
    ok = check_pw(body.password, u["pw_hash"] if u else "pbkdf2$00$00") and u and u["active"]
    if not ok:
        _fails[key] = recent + [now]
        raise HTTPException(401, "username અથવા password ખોટો છે")
    _fails.pop(key, None)
    return {"token": make_token(u["id"]), "user": public_user(u)}

@app.get("/api/me")
def me(u=Depends(current_user)):
    return public_user(u)

@app.post("/api/me/password")
def change_password(body: PwChange, u=Depends(current_user)):
    with db() as con:
        row = con.execute("SELECT pw_hash FROM staff WHERE id=?", (u["id"],)).fetchone()
        if not check_pw(body.old_password, row["pw_hash"]):
            raise HTTPException(403, "જૂનો password ખોટો છે")
        con.execute("UPDATE staff SET pw_hash=? WHERE id=?", (hash_pw(body.new_password), u["id"]))
    return {"ok": True}

# ---------- admin: staff management ----------
def _set_batches(con, uid, ids):
    ids = sorted(set(ids))
    for i in ids:
        if not con.execute("SELECT 1 FROM batches WHERE id=?", (i,)).fetchone():
            raise HTTPException(422, "પસંદ કરેલ બેચ મળ્યો નથી")
    con.execute("DELETE FROM staff_batches WHERE staff_id=?", (uid,))
    con.executemany("INSERT INTO staff_batches(staff_id,batch_id) VALUES(?,?)", [(uid, i) for i in ids])

@app.get("/api/users")
def list_users(_=Depends(need("admin"))):
    with db() as con:
        rows = con.execute("SELECT id,username,name,role,subject,active FROM staff ORDER BY role,name").fetchall()
        bm = {}
        for r in con.execute("SELECT staff_id,batch_id FROM staff_batches").fetchall():
            bm.setdefault(r["staff_id"], []).append(r["batch_id"])
    return [{**r, "active": bool(r["active"]), "batch_ids": bm.get(r["id"], [])} for r in rows]

@app.post("/api/users")
def create_user(body: UserIn, _=Depends(need("admin"))):
    with db() as con:
        if con.execute("SELECT 1 FROM staff WHERE username=?", (body.username,)).fetchone():
            raise HTTPException(409, "આ username પહેલેથી છે")
        con.execute("INSERT INTO staff(username,name,role,subject,pw_hash,active,created_at) VALUES(?,?,?,?,?,1,?)",
                    (body.username, body.name.strip(), body.role, body.subject.strip(), hash_pw(body.password), now_ist()))
        uid = con.execute("SELECT id FROM staff WHERE username=?", (body.username,)).fetchone()["id"]
        _set_batches(con, uid, body.batch_ids)
    return {"ok": True}

@app.put("/api/users/{uid}")
def update_user(uid: int, body: UserUpd, _=Depends(need("admin"))):
    with db() as con:
        t = con.execute("SELECT * FROM staff WHERE id=?", (uid,)).fetchone()
        if not t:
            raise HTTPException(404, "સ્ટાફ મળ્યો નથી")
        if t["role"] == "admin" and (body.role is not None or body.active is False):
            raise HTTPException(422, "એડમિનની ભૂમિકા કે સ્થિતિ બદલી શકાતી નથી")
        sets, vals = [], []
        if body.name is not None: sets.append("name=?"); vals.append(body.name.strip())
        if body.subject is not None: sets.append("subject=?"); vals.append(body.subject.strip())
        if body.role is not None: sets.append("role=?"); vals.append(body.role)
        if body.active is not None: sets.append("active=?"); vals.append(1 if body.active else 0)
        if body.password: sets.append("pw_hash=?"); vals.append(hash_pw(body.password))
        if sets:
            con.execute(f"UPDATE staff SET {', '.join(sets)} WHERE id=?", (*vals, uid))
        if body.batch_ids is not None and t["role"] != "admin":
            _set_batches(con, uid, body.batch_ids)
    return {"ok": True}

# ---------- batches ----------
@app.get("/api/public/batches")
def public_batches():
    with db() as con:
        return con.execute("SELECT id,name FROM batches WHERE active=1 ORDER BY name").fetchall()

@app.get("/api/batches")
def my_batches(u=Depends(current_user)):
    cnt = "(SELECT COUNT(*) FROM students s WHERE s.batch_id=b.id AND s.active=1) AS student_count"
    with db() as con:
        if u["role"] == "admin":
            rows = con.execute(f"SELECT b.id,b.name,b.description,b.active,{cnt} FROM batches b ORDER BY b.active DESC,b.name").fetchall()
        else:
            rows = con.execute(f"""SELECT b.id,b.name,b.description,b.active,{cnt} FROM batches b
                                   JOIN staff_batches sb ON sb.batch_id=b.id AND sb.staff_id=?
                                   WHERE b.active=1 ORDER BY b.name""", (u["id"],)).fetchall()
    return [{**r, "active": bool(r["active"])} for r in rows]

@app.post("/api/batches")
def create_batch(body: BatchIn, _=Depends(need("admin"))):
    with db() as con:
        if con.execute("SELECT 1 FROM batches WHERE name=?", (body.name.strip(),)).fetchone():
            raise HTTPException(409, "આ નામનો બેચ પહેલેથી છે")
        con.execute("INSERT INTO batches(name,description,active,created_at) VALUES(?,?,1,?)",
                    (body.name.strip(), body.description.strip(), now_ist()))
        bid = con.execute("SELECT id FROM batches WHERE name=?", (body.name.strip(),)).fetchone()["id"]
    return {"ok": True, "id": bid}

@app.put("/api/batches/{bid}")
def update_batch(bid: int, body: BatchUpd, _=Depends(need("admin"))):
    with db() as con:
        if not con.execute("SELECT 1 FROM batches WHERE id=?", (bid,)).fetchone():
            raise HTTPException(404, "બેચ મળ્યો નથી")
        if body.name is not None:
            dup = con.execute("SELECT id FROM batches WHERE name=?", (body.name.strip(),)).fetchone()
            if dup and dup["id"] != bid:
                raise HTTPException(409, "આ નામનો બેચ પહેલેથી છે")
        sets, vals = [], []
        if body.name is not None: sets.append("name=?"); vals.append(body.name.strip())
        if body.description is not None: sets.append("description=?"); vals.append(body.description.strip())
        if body.active is not None: sets.append("active=?"); vals.append(1 if body.active else 0)
        if sets:
            con.execute(f"UPDATE batches SET {', '.join(sets)} WHERE id=?", (*vals, bid))
    return {"ok": True}

# ---------- students (admin manages; staff read roster) ----------
@app.get("/api/batches/{bid}/students")
def batch_students(bid: int, u=Depends(need("admin", "faculty", "library"))):
    with db() as con:
        batch_access(con, u, bid)
        if u["role"] == "admin":
            rows = con.execute("SELECT id,roll,name,college,mobile,parent_name,parent_mobile,active FROM students "
                               "WHERE batch_id=? ORDER BY roll", (bid,)).fetchall()
            return [{**r, "active": bool(r["active"])} for r in rows]
        return con.execute("SELECT id,roll,name,college FROM students WHERE batch_id=? AND active=1 ORDER BY roll", (bid,)).fetchall()

def _next_roll(con, bid):
    return con.execute("SELECT COALESCE(MAX(roll),0)+1 AS n FROM students WHERE batch_id=?", (bid,)).fetchone()["n"]

@app.post("/api/students")
def create_student(body: StudentIn, _=Depends(need("admin"))):
    with db() as con:
        if not con.execute("SELECT 1 FROM batches WHERE id=?", (body.batch_id,)).fetchone():
            raise HTTPException(404, "બેચ મળ્યો નથી")
        roll = body.roll or _next_roll(con, body.batch_id)
        if con.execute("SELECT 1 FROM students WHERE batch_id=? AND roll=?", (body.batch_id, roll)).fetchone():
            raise HTTPException(409, f"Roll {roll} આ બેચમાં પહેલેથી છે")
        con.execute("INSERT INTO students(batch_id,roll,name,college,mobile,parent_name,parent_mobile,active,created_at) "
                    "VALUES(?,?,?,?,?,?,?,1,?)", (body.batch_id, roll, body.name.strip(), body.college.strip(),
                    body.mobile, body.parent_name.strip(), body.parent_mobile, now_ist()))
    return {"ok": True, "roll": roll}

@app.put("/api/students/{sid}")
def update_student(sid: int, body: StudentUpd, _=Depends(need("admin"))):
    with db() as con:
        st = con.execute("SELECT * FROM students WHERE id=?", (sid,)).fetchone()
        if not st:
            raise HTTPException(404, "વિદ્યાર્થી મળ્યો નથી")
        if body.roll is not None and body.roll != st["roll"] and con.execute(
                "SELECT 1 FROM students WHERE batch_id=? AND roll=?", (st["batch_id"], body.roll)).fetchone():
            raise HTTPException(409, f"Roll {body.roll} આ બેચમાં પહેલેથી છે")
        sets, vals = [], []
        for f in ("roll", "name", "college", "parent_name"):
            v = getattr(body, f)
            if v is not None: sets.append(f"{f}=?"); vals.append(v.strip() if isinstance(v, str) else v)
        for f in ("mobile", "parent_mobile"):
            v = getattr(body, f)
            if v is not None:
                try: sets.append(f"{f}=?"); vals.append(_chk_phone(v))
                except ValueError as e: raise HTTPException(422, str(e))
        if body.active is not None: sets.append("active=?"); vals.append(1 if body.active else 0)
        if sets:
            con.execute(f"UPDATE students SET {', '.join(sets)} WHERE id=?", (*vals, sid))
    return {"ok": True}

@app.post("/api/students/bulk")
def bulk_students(body: BulkIn, _=Depends(need("admin"))):
    errors, clean = [], []
    with db() as con:
        if not con.execute("SELECT 1 FROM batches WHERE id=?", (body.batch_id,)).fetchone():
            raise HTTPException(404, "બેચ મળ્યો નથી")
        used = {r["roll"] for r in con.execute("SELECT roll FROM students WHERE batch_id=?", (body.batch_id,)).fetchall()}
        nxt = max(used, default=0) + 1
        for i, r in enumerate(body.rows, 1):
            name = r.name.strip()
            if not name: errors.append(f"લાઇન {i}: નામ ખૂટે છે"); continue
            try: mobile = _chk_phone(r.mobile)
            except ValueError: errors.append(f"લાઇન {i}: મોબાઇલ નંબર ખોટો છે"); continue
            roll = r.roll or nxt
            if roll in used: errors.append(f"લાઇન {i}: Roll {roll} પહેલેથી છે"); continue
            used.add(roll); nxt = max(nxt, roll + 1)
            clean.append((body.batch_id, roll, name, r.college.strip(), mobile, now_ist()))
        if errors:
            raise HTTPException(422, "; ".join(errors[:6]) + (f" … (કુલ {len(errors)} ભૂલ)" if len(errors) > 6 else ""))
        con.executemany("INSERT INTO students(batch_id,roll,name,college,mobile,active,created_at) VALUES(?,?,?,?,?,1,?)", clean)
    return {"ok": True, "created": len(clean)}

# ---------- attendance sessions ----------
def visible_slots(u):
    return ("L1", "L2", "LIB") if u["role"] == "admin" else ("LIB",) if u["role"] == "library" else ("L1", "L2")

def _check_staff(con, fid, slot):
    f = con.execute("SELECT role,active FROM staff WHERE id=?", (fid,)).fetchone()
    if not f or not f["active"] or f["role"] not in SLOT_ROLES[slot]:
        raise HTTPException(422, "પસંદ કરેલ સ્ટાફ આ સત્ર માટે યોગ્ય નથી")

@app.get("/api/attendance/day")
def day_overview(batch_id: int, date: str, u=Depends(need("admin", "faculty", "library"))):
    try: _chk_date(date)
    except ValueError as e: raise HTTPException(422, str(e))
    with db() as con:
        batch_access(con, u, batch_id)
        rows = con.execute("""SELECT s.slot,s.subject,s.topic,f.name AS faculty FROM sessions s
                              LEFT JOIN staff f ON f.id=s.faculty_id WHERE s.batch_id=? AND s.date=? ORDER BY s.slot""",
                           (batch_id, date)).fetchall()
    return [r for r in rows if r["slot"] in visible_slots(u)]

@app.get("/api/attendance/recent")
def recent_sessions(batch_id: int, u=Depends(need("admin", "faculty", "library"))):
    slots = visible_slots(u)
    with db() as con:
        batch_access(con, u, batch_id)
        rows = con.execute("""SELECT s.date,s.slot,s.subject,s.topic,f.name AS faculty,
            (SELECT COUNT(*) FROM att a WHERE a.batch_id=s.batch_id AND a.date=s.date AND a.slot=s.slot AND a.status='Present') AS present,
            (SELECT COUNT(*) FROM att a WHERE a.batch_id=s.batch_id AND a.date=s.date AND a.slot=s.slot) AS total
            FROM sessions s LEFT JOIN staff f ON f.id=s.faculty_id WHERE s.batch_id=?
            ORDER BY s.date DESC, s.slot LIMIT 150""", (batch_id,)).fetchall()
    return [r for r in rows if r["slot"] in slots][:40]

@app.get("/api/attendance/session")
def get_session(batch_id: int, date: str, slot: Slot, u=Depends(need("admin", "faculty", "library"))):
    try: _chk_date(date)
    except ValueError as e: raise HTTPException(422, str(e))
    if slot not in visible_slots(u):
        raise HTTPException(403, "આ સત્રની વિગત જોવાની તમારી પરવાનગી નથી")
    with db() as con:
        batch_access(con, u, batch_id)
        s = con.execute("""SELECT s.faculty_id,s.subject,s.topic,f.name AS faculty FROM sessions s
                           LEFT JOIN staff f ON f.id=s.faculty_id WHERE s.batch_id=? AND s.date=? AND s.slot=?""",
                        (batch_id, date, slot)).fetchone()
        rows = con.execute("""SELECT a.student_id,st.roll,st.name,st.college,a.status,a.lib_in,a.lib_out,a.book
                              FROM att a JOIN students st ON st.id=a.student_id
                              WHERE a.batch_id=? AND a.date=? AND a.slot=? ORDER BY st.roll""", (batch_id, date, slot)).fetchall()
    allowed = u["role"] in SLOT_ROLES[slot]
    return {"exists": bool(s), "session": s, "rows": rows, "allowed": allowed,
            "can_edit": allowed and (u["role"] == "admin" or not s)}

@app.put("/api/attendance/session")
def save_session(body: SessionIn, u=Depends(need("admin", "faculty", "library"))):
    if u["role"] not in SLOT_ROLES[body.slot]:
        raise HTTPException(403, "આ સત્રની એન્ટ્રી કરવાની તમારી પરવાનગી નથી")
    if u["role"] != "admin" and body.date > today_ist():
        raise HTTPException(422, "ભવિષ્યની તારીખની એન્ટ્રી થઈ શકતી નથી")
    ids = [r.student_id for r in body.rows]
    if not ids or len(set(ids)) != len(ids):
        raise HTTPException(422, "વિદ્યાર્થી યાદી ખાલી છે અથવા ડુપ્લિકેટ છે")
    with db() as con:
        batch_access(con, u, body.batch_id)
        stu = {r["id"]: r for r in con.execute("SELECT id,roll,active FROM students WHERE batch_id=?", (body.batch_id,)).fetchall()}
        if any(i not in stu for i in ids):
            raise HTTPException(422, "કોઈ વિદ્યાર્થી આ બેચનો નથી")
        if body.slot == "LIB":
            for r in body.rows:
                if r.status == "Present":
                    roll = stu[r.student_id]["roll"]
                    if r.lib_in and r.lib_in < "12:00":
                        raise HTTPException(422, f"Roll {roll}: લાયબ્રેરી In 12:00 પછી હોવો જોઈએ")
                    if r.lib_in and r.lib_out and r.lib_out <= r.lib_in:
                        raise HTTPException(422, f"Roll {roll}: લાયબ્રેરી Out સમય In પછી હોવો જોઈએ")
        ex = con.execute("SELECT faculty_id FROM sessions WHERE batch_id=? AND date=? AND slot=?",
                         (body.batch_id, body.date, body.slot)).fetchone()
        if ex and u["role"] != "admin":
            raise HTTPException(403, "આ તારીખ/સત્રની હાજરી પહેલેથી સેવ છે. ફેરફાર ફક્ત એડમિન કરી શકશે.")
        if not ex and set(ids) != {i for i, s in stu.items() if s["active"]}:
            raise HTTPException(422, "વિદ્યાર્થી યાદી બદલાઈ છે. પેજ રિફ્રેશ કરીને ફરી પ્રયત્ન કરો.")
        fid = ex["faculty_id"] if ex else u["id"]
        if u["role"] == "admin" and body.faculty_id:
            _check_staff(con, body.faculty_id, body.slot)
            fid = body.faculty_id
        key = (body.batch_id, body.date, body.slot)
        try:
            if ex:
                con.execute("UPDATE sessions SET faculty_id=?,subject=?,topic=?,updated_at=? WHERE batch_id=? AND date=? AND slot=?",
                            (fid, body.subject.strip(), body.topic.strip(), now_ist(), *key))
                con.execute("DELETE FROM att WHERE batch_id=? AND date=? AND slot=?", key)
            else:
                con.execute("INSERT INTO sessions(batch_id,date,slot,faculty_id,subject,topic,updated_at) VALUES(?,?,?,?,?,?,?)",
                            (*key, fid, body.subject.strip(), body.topic.strip(), now_ist()))
            lib = body.slot == "LIB"
            con.executemany("INSERT INTO att(batch_id,date,slot,student_id,status,lib_in,lib_out,book) VALUES(?,?,?,?,?,?,?,?)",
                [(*key, r.student_id, r.status,
                  r.lib_in if lib and r.status == "Present" else None,
                  r.lib_out if lib and r.status == "Present" else None,
                  (r.book or None) if lib and r.status == "Present" else None) for r in body.rows])
        except INTEGRITY:
            raise HTTPException(409, "આ સત્રની એન્ટ્રી હમણાં જ કોઈએ સેવ કરી છે. પેજ રિફ્રેશ કરો.")
    return {"ok": True, "present": sum(r.status == "Present" for r in body.rows), "total": len(body.rows)}

@app.delete("/api/attendance/session")
def delete_session(batch_id: int, date: str, slot: Slot, _=Depends(need("admin"))):
    with db() as con:
        con.execute("DELETE FROM att WHERE batch_id=? AND date=? AND slot=?", (batch_id, date, slot))
        con.execute("DELETE FROM sessions WHERE batch_id=? AND date=? AND slot=?", (batch_id, date, slot))
    return {"ok": True}

# ---------- tests ----------
@app.get("/api/tests/list")
def tests_list(batch_id: int, u=Depends(need("admin", "faculty"))):
    with db() as con:
        batch_access(con, u, batch_id)
        return con.execute("""SELECT t.id,t.date,t.title,t.max_marks,f.name AS faculty FROM tests t
                              LEFT JOIN staff f ON f.id=t.created_by WHERE t.batch_id=? ORDER BY t.date DESC,t.id DESC""",
                           (batch_id,)).fetchall()

@app.get("/api/tests/report")
def tests_report(batch_id: int, month: Optional[str] = Query(None, pattern=MONTH_RE), u=Depends(need("admin", "faculty"))):
    with db() as con:
        batch_access(con, u, batch_id)
        tests, rows = build_test_report(con, batch_id, month)
    return {"month": month, "tests": tests, "rows": rows}

@app.get("/api/tests/{tid}")
def test_detail(tid: int, u=Depends(need("admin", "faculty"))):
    with db() as con:
        t = con.execute("""SELECT t.id,t.batch_id,t.date,t.title,t.max_marks,t.created_by,f.name AS faculty FROM tests t
                           LEFT JOIN staff f ON f.id=t.created_by WHERE t.id=?""", (tid,)).fetchone()
        if not t: raise HTTPException(404, "ટેસ્ટ મળ્યો નથી")
        batch_access(con, u, t["batch_id"])
        marks = con.execute("""SELECT m.student_id,st.roll,st.name,st.college,m.score FROM marks m
                               JOIN students st ON st.id=m.student_id WHERE m.test_id=? ORDER BY st.roll""", (tid,)).fetchall()
    return {"test": t, "marks": marks, "can_edit": u["role"] == "admin"}

@app.put("/api/tests")
def save_test(body: TestIn, u=Depends(need("admin", "faculty"))):
    if u["role"] != "admin" and (body.test_id or body.date > today_ist()):
        raise HTTPException(403 if body.test_id else 422,
                            "ફેરફાર ફક્ત એડમિન કરી શકશે." if body.test_id else "ભવિષ્યની તારીખનો ટેસ્ટ સેવ થઈ શકતો નથી")
    ids = [m.student_id for m in body.marks]
    if not ids or len(set(ids)) != len(ids):
        raise HTTPException(422, "વિદ્યાર્થી યાદી ખાલી છે અથવા ડુપ્લિકેટ છે")
    title = body.title.strip()
    with db() as con:
        old = None
        bid = body.batch_id
        if body.test_id:
            old = con.execute("SELECT batch_id,created_by FROM tests WHERE id=?", (body.test_id,)).fetchone()
            if not old: raise HTTPException(404, "ટેસ્ટ મળ્યો નથી")
            bid = old["batch_id"]
        batch_access(con, u, bid)
        stu = {r["id"]: r for r in con.execute("SELECT id,roll,active FROM students WHERE batch_id=?", (bid,)).fetchall()}
        if any(i not in stu for i in ids):
            raise HTTPException(422, "કોઈ વિદ્યાર્થી આ બેચનો નથી")
        if not old and set(ids) != {i for i, s in stu.items() if s["active"]}:
            raise HTTPException(422, "વિદ્યાર્થી યાદી બદલાઈ છે. પેજ રિફ્રેશ કરીને ફરી પ્રયત્ન કરો.")
        for m in body.marks:
            if m.score is not None and not (0 <= m.score <= body.max_marks):
                raise HTTPException(422, f"Roll {stu[m.student_id]['roll']}: ગુણ 0 થી {body.max_marks:g} વચ્ચે હોવા જોઈએ")
        dup = con.execute("SELECT id FROM tests WHERE batch_id=? AND date=? AND title=?", (bid, body.date, title)).fetchone()
        if dup and dup["id"] != body.test_id:
            raise HTTPException(409, "આ તારીખે આ નામનો ટેસ્ટ પહેલેથી છે. ફેરફાર માટે એડમિનને કહો.")
        owner = u["id"]
        if u["role"] == "admin" and body.faculty_id:
            _check_staff(con, body.faculty_id, "L1")
            owner = body.faculty_id
        if old:
            tid = body.test_id
            con.execute("UPDATE tests SET date=?,title=?,max_marks=?,created_by=? WHERE id=?",
                        (body.date, title, body.max_marks, body.faculty_id or old["created_by"], tid))
            con.execute("DELETE FROM marks WHERE test_id=?", (tid,))
        else:
            con.execute("INSERT INTO tests(batch_id,date,title,max_marks,created_by) VALUES(?,?,?,?,?)",
                        (bid, body.date, title, body.max_marks, owner))
            tid = con.execute("SELECT id FROM tests WHERE batch_id=? AND date=? AND title=?", (bid, body.date, title)).fetchone()["id"]
        con.executemany("INSERT INTO marks(test_id,student_id,score) VALUES(?,?,?)", [(tid, m.student_id, m.score) for m in body.marks])
    return {"ok": True, "test_id": tid}

@app.delete("/api/tests/{tid}")
def delete_test(tid: int, _=Depends(need("admin"))):
    with db() as con:
        con.execute("DELETE FROM marks WHERE test_id=?", (tid,))
        con.execute("DELETE FROM tests WHERE id=?", (tid,))
    return {"ok": True}

# ---------- public: student lookup (no mobile numbers are exposed) ----------
def student_detail(con, st):
    batch = con.execute("SELECT name FROM batches WHERE id=?", (st["batch_id"],)).fetchone()
    recs = con.execute("""SELECT a.date,a.slot,a.status,a.lib_in,a.lib_out,a.book,s.subject,s.topic,f.name AS faculty
                          FROM att a JOIN sessions s ON s.batch_id=a.batch_id AND s.date=a.date AND s.slot=a.slot
                          LEFT JOIN staff f ON f.id=s.faculty_id WHERE a.student_id=?""", (st["id"],)).fetchall()
    days, held, attended, lib_days = {}, 0, 0, 0
    for r in recs:
        e = {"status": r["status"], "faculty": r["faculty"], "subject": r["subject"], "topic": r["topic"]}
        if r["slot"] == "LIB":
            e.update(lib_in=r["lib_in"], lib_out=r["lib_out"], book=r["book"]); lib_days += r["status"] == "Present"
        else:
            held += 1; attended += r["status"] == "Present"
        days.setdefault(r["date"], {"date": r["date"]})[r["slot"]] = e
    tests, got, mx = [], 0.0, 0.0
    for t in con.execute("""SELECT t.date,t.title,t.max_marks,m.score FROM marks m JOIN tests t ON t.id=m.test_id
                            WHERE m.student_id=? ORDER BY t.date DESC,t.id DESC""", (st["id"],)).fetchall():
        t["percent"] = round(t["score"] / t["max_marks"] * 100, 1) if t["score"] is not None else None
        if t["score"] is not None: got += t["score"]; mx += t["max_marks"]
        tests.append(t)
    return {"student": {"roll": st["roll"], "name": st["name"], "college": st["college"], "batch": batch["name"]},
            "attendance": [days[d] for d in sorted(days, reverse=True)],
            "summary": {"lectures_held": held, "lectures_attended": attended, "library_days": lib_days,
                        "percent": round(attended / held * 100, 1) if held else 0.0,
                        "test_percent": round(got / mx * 100, 1) if mx else None,
                        "tests_taken": sum(t["score"] is not None for t in tests)},
            "tests": tests}

@app.get("/api/student")
def find_student(batch_id: int, q: str = Query(min_length=1, max_length=60)):
    q = q.strip()
    with db() as con:
        if not con.execute("SELECT 1 FROM batches WHERE id=? AND active=1", (batch_id,)).fetchone():
            raise HTTPException(404, "બેચ મળ્યો નથી")
        if q.isdigit():
            st = con.execute("SELECT * FROM students WHERE batch_id=? AND roll=? AND active=1", (batch_id, int(q))).fetchone()
            if not st: raise HTTPException(404, "આ Roll નો વિદ્યાર્થી આ બેચમાં મળ્યો નથી")
            return student_detail(con, st)
        found = con.execute("SELECT * FROM students WHERE batch_id=? AND active=1 AND lower(name) LIKE ? ORDER BY roll",
                            (batch_id, f"%{q.lower()}%")).fetchall()
        if not found: raise HTTPException(404, "વિદ્યાર્થી મળ્યો નથી")
        if len(found) > 1: return {"matches": [{"roll": s["roll"], "name": s["name"]} for s in found]}
        return student_detail(con, found[0])

# ---------- reports ----------
def build_report(con, batch_id, month):
    like = month + "-%"
    keys = [(s["date"], s["slot"]) for s in con.execute(
        "SELECT date,slot FROM sessions WHERE batch_id=? AND date LIKE ? AND slot IN ('L1','L2') ORDER BY date,slot",
        (batch_id, like)).fetchall()]
    status, lib = {}, Counter()
    for r in con.execute("SELECT date,slot,student_id,status FROM att WHERE batch_id=? AND date LIKE ?", (batch_id, like)).fetchall():
        status[(r["date"], r["slot"], r["student_id"])] = r["status"]
        if r["slot"] == "LIB" and r["status"] == "Present": lib[r["student_id"]] += 1
    rows = []
    for s in con.execute("SELECT * FROM students WHERE batch_id=? ORDER BY roll", (batch_id,)).fetchall():
        recs = [status.get((d, sl, s["id"])) for d, sl in keys]
        held, p = sum(1 for x in recs if x), sum(1 for x in recs if x == "Present")
        if not s["active"] and not held and not lib[s["id"]]:
            continue
        rows.append({"student_id": s["id"], "roll": s["roll"], "name": s["name"], "college": s["college"],
                     "held": held, "present": p, "absent": held - p, "library_days": lib[s["id"]],
                     "percent": round(p / held * 100, 1) if held else 0.0})
    return keys, status, rows

def build_test_report(con, batch_id, month=None):
    q = ("SELECT t.id,t.date,t.title,t.max_marks,f.name AS faculty FROM tests t "
         "LEFT JOIN staff f ON f.id=t.created_by WHERE t.batch_id=?")
    tests = (con.execute(q + " AND t.date LIKE ? ORDER BY t.date,t.id", (batch_id, month + "-%")) if month
             else con.execute(q + " ORDER BY t.date,t.id", (batch_id,))).fetchall()
    ids = {t["id"] for t in tests}
    score = {(m["test_id"], m["student_id"]): m["score"] for m in con.execute(
        "SELECT m.test_id,m.student_id,m.score FROM marks m JOIN tests t ON t.id=m.test_id WHERE t.batch_id=?", (batch_id,)).fetchall()
        if m["test_id"] in ids}
    rows = []
    for s in con.execute("SELECT * FROM students WHERE batch_id=? ORDER BY roll", (batch_id,)).fetchall():
        sc, got, mx, ap = [], 0.0, 0.0, 0
        for t in tests:
            v = score.get((t["id"], s["id"])); sc.append(v)
            if v is not None: got += v; mx += t["max_marks"]; ap += 1
        if not s["active"] and not ap:
            continue
        rows.append({"student_id": s["id"], "roll": s["roll"], "name": s["name"], "college": s["college"],
                     "scores": sc, "appeared": ap, "total_obtained": got, "total_max": mx,
                     "percent": round(got / mx * 100, 1) if mx else None})
    order = sorted({r["percent"] for r in rows if r["percent"] is not None}, reverse=True)
    for r in rows:
        r["rank"] = order.index(r["percent"]) + 1 if r["percent"] is not None else None
    return tests, rows

@app.get("/api/report")
def report(batch_id: int, month: str = Query(pattern=MONTH_RE), u=Depends(need("admin", "faculty"))):
    with db() as con:
        batch_access(con, u, batch_id)
        keys, _st, rows = build_report(con, batch_id, month)
    return {"month": month, "lectures_held": len(keys), "rows": rows}

# ---------- Excel export ----------
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HEAD_FILL = PatternFill("solid", fgColor="3B1E08")
GREEN, RED = PatternFill("solid", fgColor="E8F5E9"), PatternFill("solid", fgColor="FFEBEE")
THIN = Side(style="thin", color="D8CBBE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def _cell(ws, row, col, value, fill=None, bold=False, color=None, wrap=False, align="center"):
    c = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        c.data_type = "s"          # typed text never becomes an Excel formula
    c.border = BORDER
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill: c.fill = fill
    if bold or color: c.font = Font(bold=bold, color=color)
    return c

def _header(ws, headers):
    for i, h in enumerate(headers, 1):
        _cell(ws, 1, i, h, fill=HEAD_FILL, bold=True, color="FFFFFF", wrap=True)
    ws.row_dimensions[1].height = 44
    ws.freeze_panes = "D2"

def _widths(ws):
    for col in ws.columns:
        ln = max((len(str(c.value).split("\n")[0]) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(8, ln + 2), 42)

def _xlsx(wb, name):
    buf = BytesIO(); wb.save(buf)
    return Response(buf.getvalue(), media_type=XLSX,
                    headers={"Content-Disposition": f'attachment; filename="{name}"', "Cache-Control": "no-store"})

@app.get("/api/export/report.xlsx")
def export_report(batch_id: int, month: str = Query(pattern=MONTH_RE), u=Depends(need("admin", "faculty"))):
    with db() as con:
        b = batch_access(con, u, batch_id)
        keys, status, rows = build_report(con, batch_id, month)
    if not keys:
        raise HTTPException(404, "આ મહિના માટે કોઈ લેક્ચર હાજરી ડેટા નથી")
    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    _header(ws, ["Roll", "Student Name", "College", "Lectures Held", "Attended", "Absent", "Attendance %", "Library Days"])
    for i, r in enumerate(rows, 2):
        _cell(ws, i, 1, r["roll"]); _cell(ws, i, 2, r["name"], align="left"); _cell(ws, i, 3, r["college"], align="left")
        _cell(ws, i, 4, r["held"]); _cell(ws, i, 5, r["present"], color="2E7D32", bold=True)
        _cell(ws, i, 6, r["absent"], color="D32F2F", bold=True)
        _cell(ws, i, 7, r["percent"], bold=True).number_format = '0.0"%"'
        _cell(ws, i, 8, r["library_days"])
    _widths(ws)
    ws.cell(row=len(rows) + 3, column=2, value=f"Batch: {b['name']}  |  Month: {month}")
    wd = wb.create_sheet("Daily")
    _header(wd, ["Roll", "Student Name", "College"] + [f"{d[8:10]}-{d[5:7]}\n{sl}" for d, sl in keys])
    for i, r in enumerate(rows, 2):
        _cell(wd, i, 1, r["roll"]); _cell(wd, i, 2, r["name"], align="left"); _cell(wd, i, 3, r["college"], align="left")
        for j, (d, sl) in enumerate(keys, 4):
            st = status.get((d, sl, r["student_id"]))
            if st is None: _cell(wd, i, j, "-")
            else:
                p = st == "Present"
                _cell(wd, i, j, "P" if p else "A", fill=GREEN if p else RED, bold=True, color="2E7D32" if p else "D32F2F")
    _widths(wd)
    for j in range(4, 4 + len(keys)):
        wd.column_dimensions[get_column_letter(j)].width = 7
    return _xlsx(wb, f"attendance_{batch_id}_{month}.xlsx")

@app.get("/api/export/tests.xlsx")
def export_tests(batch_id: int, month: Optional[str] = Query(None, pattern=MONTH_RE), u=Depends(need("admin", "faculty"))):
    with db() as con:
        b = batch_access(con, u, batch_id)
        tests, rows = build_test_report(con, batch_id, month)
    if not tests:
        raise HTTPException(404, "કોઈ ટેસ્ટ ડેટા નથી")
    wb = Workbook(); ws = wb.active; ws.title = "Test Marks"
    heads = ["Roll", "Student Name", "College"] + [
        f"{t['title']}\n{t['date']} (/{t['max_marks']:g})\n{t['faculty'] or ''}" for t in tests]
    _header(ws, heads + ["Total Obtained", "Total Max", "Percentage", "Rank"])
    ws.row_dimensions[1].height = 60
    for i, r in enumerate(rows, 2):
        _cell(ws, i, 1, r["roll"]); _cell(ws, i, 2, r["name"], align="left"); _cell(ws, i, 3, r["college"], align="left")
        for j, v in enumerate(r["scores"], 4):
            if v is None: _cell(ws, i, j, "AB", fill=RED, bold=True, color="D32F2F")
            else: _cell(ws, i, j, v)
        k = 4 + len(tests)
        _cell(ws, i, k, r["total_obtained"] if r["appeared"] else "-")
        _cell(ws, i, k + 1, r["total_max"] if r["appeared"] else "-")
        c = _cell(ws, i, k + 2, r["percent"] if r["percent"] is not None else "-", bold=True)
        if r["percent"] is not None: c.number_format = '0.0"%"'
        _cell(ws, i, k + 3, r["rank"] if r["rank"] else "-")
    _widths(ws)
    for j in range(4, 4 + len(tests)):
        ws.column_dimensions[get_column_letter(j)].width = 20
    ws.cell(row=len(rows) + 3, column=2, value=f"Batch: {b['name']}  |  AB = Absent. Percentage = total obtained / total marks of tests attended.")
    return _xlsx(wb, f"test_results_{batch_id}_{month or 'all'}.xlsx")

# ---------- health, page, PWA assets ----------
@app.get("/api/health")
def health():
    return {"ok": True}      # no DB call -> an uptime pinger keeps the server awake cheaply

ASSETS = {"manifest.webmanifest": "application/manifest+json", "sw.js": "application/javascript",
          "icon-192.png": "image/png", "icon-512.png": "image/png", "logo.png": "image/png"}

def _find(name):
    for d in (BASE / "static", BASE):
        if (d / name).exists():
            return d / name

@app.get("/")
def index():
    p = _find("index.html")
    if not p:
        raise HTTPException(500, "index.html મળ્યું નથી. main.py ની બાજુમાં (અથવા static ફોલ્ડરમાં) મૂકો.")
    return FileResponse(p, headers={"Cache-Control": "no-cache"})

@app.get("/{name}")
def asset(name: str):
    p = _find(name) if name in ASSETS else None
    if not p:
        raise HTTPException(404, "Not found")
    return FileResponse(p, media_type=ASSETS[name], headers={"Cache-Control": "no-cache"})
