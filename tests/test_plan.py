# Run: python3 tests/test_plan.py  (no dependencies needed)
import importlib.util, sys
from datetime import datetime, timedelta, timezone, date
spec = importlib.util.spec_from_file_location("calsync", str(__import__("pathlib").Path(__file__).resolve().parent.parent / "calsync.py"))
m = importlib.util.module_from_spec(spec); sys.modules["calsync"]=m; spec.loader.exec_module(m)
E = m.Event
cfg = {"sync": dict(m.DEFAULTS)}
now = datetime(2026,11,9,21,tzinfo=timezone.utc)
old = now - timedelta(hours=1)
t = lambda h: datetime(2026,11,10,h,tzinfo=timezone.utc)
def ev(side,id,title,h,mark=None,upd=old,**kw): return E(side,id,title,False,t(h),t(h+1),upd,mark,**kw)
def kinds(a): return [(x[0], x[1].id) for x in a]

# 1. plain copy
assert kinds(m.plan("google",[ev("google","g1","Lab mtg",9)],[],cfg,now))==[("create","g1")]
# 2. unsettled -> wait
assert m.plan("google",[ev("google","g1","Lab mtg",9,upd=now)],[],cfg,now)==[]
# 3. native exists with prefix-match title -> no copy
assert m.plan("google",[ev("google","g1","Lab meeting",9)],[ev("outlook","o1","Lab meeting - room 3",9)],cfg,now)==[]
# 4. our copy + native appeared -> delete ours
a=m.plan("google",[ev("google","g1","Lab meeting",9)],[ev("outlook","o1","Lab meeting",9),ev("outlook","c1","Lab meeting",9,mark="google:g1")],cfg,now)
assert kinds(a)==[("delete","c1")],a
# 5. our copies never become sources (no loop)
assert m.plan("outlook",[ev("outlook","c1","X meeting",9,mark="google:g1")],[ev("google","g1","X meeting",9)],cfg,now)==[]
# 6. source moved -> update
a=m.plan("google",[ev("google","g1","Lab meeting",10)],[ev("outlook","c1","Lab meeting",9,mark="google:g1")],cfg,now)
assert kinds(a)==[("update","c1")]
# 7. source deleted -> delete copy; copies from other side untouched
a=m.plan("google",[],[ev("outlook","c1","A",9,mark="google:g1"),ev("outlook","c2","B",9,mark="outlook:zzz")],cfg,now)
assert kinds(a)==[("delete","c1")],a
# 8. declined -> delete copy
a=m.plan("google",[ev("google","g1","A meeting",9,copyable=False)],[ev("outlook","c1","A meeting",9,mark="google:g1")],cfg,now)
assert kinds(a)==[("delete","c1")]
# 9. short title prefix doesn't over-match
assert kinds(m.plan("google",[ev("google","g1","1:1",9)],[ev("outlook","o1","1:1 planning",9)],cfg,now))==[("create","g1")]
# 10. different time -> copy
assert kinds(m.plan("google",[ev("google","g1","Lab",9)],[ev("outlook","o1","Lab",10)],cfg,now))==[("create","g1")]
# 11. private title
a=m.plan("google",[ev("google","g1","Doctor",9,private=True)],[ev("outlook","c1","Busy",9,mark="google:g1")],cfg,now)
assert a==[]
# 12. duplicate copies
a=m.plan("google",[ev("google","g1","Lab",9)],[ev("outlook","c1","Lab",9,mark="google:g1"),ev("outlook","c2","Lab",9,mark="google:g1")],cfg,now)
assert kinds(a)==[("delete","c2")]
# parse_ts with Graph 7-digit fractions
assert m.parse_ts("2026-11-10T15:00:00.0000000")==t(15)
assert m.parse_ts("2026-11-09T20:59:44.1234567Z").microsecond==123456
# outlook all-day parse (local tz America/Chicago)
import os,time; os.environ["TZ"]="America/Chicago"; time.tzset()
o=m.Outlook.__new__(m.Outlook)
e=o._parse({"id":"x","subject":"Holiday","isAllDay":True,"start":{"dateTime":"2026-11-26T06:00:00.0000000"},"end":{"dateTime":"2026-11-27T06:00:00.0000000"},"lastModifiedDateTime":"2026-11-01T00:00:00Z","showAs":"busy"})
assert (e.start,e.end)==(date(2026,11,26),date(2026,11,27)),e
print("all ok")
