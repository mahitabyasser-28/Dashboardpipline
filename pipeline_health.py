# =====================================================================
# OPEN OPPORTUNITY HEALTH PIPELINE — ALL SMs · TLs · Sales Reps
# ---------------------------------------------------------------------
# WHAT IT DOES
#   Pulls every OPEN opportunity (all open stages, read from Salesforce's
#   OpportunityStage table) and builds ONE self-contained, clickable HTML
#   file for Sales Managers and C-level: filters, stage funnel, risk
#   reasons, SM → TL → Rep breakdown with drill-down, and a full opp list.
#
# HIERARCHY (same rules as the EOD Performance Tracker):
#   · SALES_MANAGER_NAMES / TEAM_LEADER_NAMES = the CURRENT org chart.
#   · Rep → Team Leader = the OWNER's "Team for Tracking membership"
#     (prefix match, e.g. "Bernadette" ≈ "Bernadette Atfy").
#   · TL → Sales Manager = first SM up the TL's Manager chain
#     (or TL_SALES_MANAGER_OVERRIDE). A TL with no SM above them is shown
#     as "Direct to <their manager>".
#   · Opps owned by a TL / SM are shown under that TL / SM.
#   · Opps whose owner matches no team are shown under "Other owners"
#     so the totals always reconcile with Salesforce.
#
# HEALTH = WHERE THE CLIENT STANDS, FROM READING THE OPPORTUNITY NOTES
#   Each note was read in full and its client placed in one stance, with a
#   one-line summary of where the client stands. Those reviewed labels live in
#   NOTE_LABELS_FILE (note_labels.json), keyed by Opportunity Id + a fingerprint
#   of the note text:
#     HEALTHY  Ready to book  (EOI / cheque / deposit / unit chosen and paying)
#              Engaged        (two-way contact: meeting, visit, reviewing offer)
#     WATCH    Waiting on us  (waiting for a release, launch, the Alex project,
#                              or a unit / size that is not available)
#              Hesitant       (postponing, thinking, travelling, selling first,
#                              money / price concern, comparing competitors)
#     AT RISK  Unresponsive   (latest attempts unanswered, phone off)
#              Losing interest(not interested, bought elsewhere, not a buyer)
#              No clear update(empty note, or only a requirement like "2br noor")
#   A note that is NEW or CHANGED since it was reviewed gets a provisional
#   stance from the word lists below and is marked "not reviewed yet"; the
#   script also writes Notes_To_Review_<date>.csv with those notes so they can
#   be read and added to note_labels.json.
#   Activity data (days in stage, close date) is shown for context only.
#
# USAGE
#   In the same notebook/session as the EOD tracker (global `sf` =
#   simple_salesforce connection) just run this file.
#   Output: Open_Pipeline_Health_<Day><Mon>.html  (e.g. …_2ndNov.html)
#   Email / share the HTML: it opens in any browser, no login needed.
#
# REVIEWED STANCES FROM GOOGLE SHEETS (updated every morning by Claude)
#   Each daily Claude run reads every new / changed note and writes the full set of
#   reviewed stances (those not built into this script) to the "Note_Labels" tab of LABELS_SHEET_URL. When you run this
#   script in Colab it downloads that tab first, so the dashboard uses the same
#   stances as the published Claude dashboard. Order of precedence:
#     built-in labels (2026-11-02)  <  Google Sheet  <  local note_labels.json
#     (a local label only wins if it was reviewed on the same day or later).
#   The sheet must be shared as "Anyone with the link can view" — or the script
#   falls back to signing in with your Google account (Colab pop-up).
#   A note edited AFTER the morning run shows "not reviewed" until the next run.
#
# ALL NOTES IN GOOGLE SHEETS
#   Every run also writes EVERY open opportunity — full note text, stance, summary,
#   health, rep / TL / SM, stage, amount — to the "All_Notes" tab (ALL_NOTES_TAB).
#   Writing needs edit access: in Colab you sign in with your Google account; on
#   GitHub Actions a Google service account (see below) that is an Editor on the sheet.
#
# RUNNING IT
#   · Colab: run your `sf = Salesforce(...)` cell, then this file (as before).
#   · GitHub Actions / any server: no `sf` needed — set these environment variables
#       SF_USERNAME, SF_PASSWORD, SF_SECURITY_TOKEN   (SF_DOMAIN = "test" for a sandbox)
#       GOOGLE_SERVICE_ACCOUNT_JSON  = the service-account key JSON (the whole text)
#     and run:  python pipeline_health.py
#   pip install simple-salesforce gspread google-auth
# =====================================================================
import json, os, re, sys, html as _html
from datetime import datetime, date, timedelta, timezone

# ------------------------- CONFIG (change per run) -------------------------
REPORT_DATE = None                         # ◀◀ None = today (Cairo)
TZ = timezone(timedelta(hours=3))          # same fixed UTC+3 as the EOD tracker

APPLY_REGION_FILTER = True                 # Egypt only
REGION_FLD, REGION_VAL = "Lead_Region__c", "Egypt"
TEAM_FLD = "team_for_tracking_membership__c"

SALES_MANAGER_NAMES = [
    "Islam Hussein", "Mohamed Abdelaziz", "Mohamed Elghoul", "Nouran Elhennawy",
]
TEAM_LEADER_NAMES = [
    "Ayman Osama", "Bernadette Atfy", "Kareem Moghazy", "Khaled Fares",
    "Mohamed Bassel", "Mohamed Hazem", "Mohamed Karam", "Mohamed Samir",
    "Mohanad Mohamed", "Mostafa Morsy", "Mostafa Zohair", "Osama El Anany",
]
TL_SALES_MANAGER_OVERRIDE = {
    # "Mohamed Karam": "Mohamed Abdelaziz",
}

NOTE_FIELDS = ["Opportunity_Notes__c", "Description"]   # ◀◀ fields read for the client's stance
NOTE_LABELS_FILE = "note_labels.json"     # ◀◀ reviewed stances (keep next to this script)
LABELS_SHEET_URL = "https://docs.google.com/spreadsheets/d/1w6Iw8OUTLDlt6qF3YmG1QZbAFwdeDVPkY5BKBXzlKjM/edit"  # ◀◀ None = don't use the sheet
LABELS_SHEET_TAB = "Note_Labels"          # ◀◀ tab Claude writes every morning
ALL_NOTES_TAB = "All_Notes"               # ◀◀ tab this script fills with every open opp + full note (None = off)
SHEET_CELL_MAX = 49000                    # Google Sheets limit is 50,000 characters per cell

CURRENCY = "EGP"
OPEN_STAGES = None        # ◀◀ None = auto from OpportunityStage (IsClosed = false), in Salesforce order
SF_BASE_URL = None        # ◀◀ None = auto from the sf connection (opp names become links)
FILE_TAG = "Open_Pipeline_Health"
# ---------------------------------------------------------------------------

# ============== PROVISIONAL RULES (only for notes not reviewed yet) ==============
import re, html as _html

# ◀◀ STANCES: order = priority when ONE update line matches several stances (first wins)
#    health: 0 healthy · 1 watch · 2 at risk
STANCES = [
    ("lost",    "Losing interest", 2),
    ("book",    "Ready to book",   0),
    ("unresp",  "Unresponsive",    2),
    ("hesit",   "Hesitant",        1),
    ("wait",    "Waiting on us",   1),
    ("engaged", "Engaged",         0),
]
NO_NOTES = ("none", "No clear update", 2)
STANCE_HEALTH = {k: h for k, _l, h in STANCES}; STANCE_HEALTH[NO_NOTES[0]] = NO_NOTES[2]

# ◀◀ KEYWORDS — add words here when you see a note read wrongly. Lower-case.
#    Arabic is matched after normalising (أ/إ/آ→ا · ة→ه · ى→ي · no tashkeel).
#    A word may carry a prefix (و ف ب ل) and any ending, e.g. "مهتم" also matches "ومهتمه".
#    English / Franco words match at the start of a word: "interest" matches "interested".
STANCE_KEYWORDS = {
    "lost": [
        "اشتري في", "اشتري من", "اشترت في", "اشترت من", "اشتري خلاص", "mshtry", "msh hshtry", "msh hashtry", "mesh hashtry", "m4 hashtry", "msh hyshtry",
        "not interested", "not intrested", "not intersted", "no interest", "lost interest", "isn't interested",
        "not serious", "wrong number", "bought from", "bought in", "bought a", "already bought", "purchased",
        "went with", "booked with another", "booked elsewhere", "no budget", "refused", "doesn't want",
        "doesnt want", "don't want", "dont want", "not looking", "changed his mind", "changed her mind", "not relevant",
        "msh mohtam", "mesh mohtam", "mosh mohtam", "mish mohtam", "msh mehtam", "mesh mehtam", "msh 3ayez", "msh 3ayz",
        "mesh 3ayez", "mesh 3ayz", "mosh 3ayez", "msh 3awz", "mesh 3awez", "eshtara", "ishtara", "eshtaret", "e4tara",
        "msh hay7gez", "mesh hay7gez", "lagha", "msh serious", "mesh serious",
        "مش مهتم", "غير مهتم", "مش مهتمه", "مبقاش مهتم", "مش عايز", "مش عاوز", "مش عايزه", "مش عاوزه", "مش حابب",
        "مش هيحجز", "مش هتحجز", "اشترت", "حجز في مكان", "حجز مع شركه", "حجزت مع", "الغي", "لغي",
        "مش جاد", "رقم غلط", "الرقم غلط", "مفيش ميزانيه", "ملوش ميزانيه", "معندوش ميزانيه", "رفض", "رفضت",
        "صرف نظر", "صرفت نظر", "مش بيدور", "مش بتدور", "مبقاش عايز",
    ],
    "book": [
        "reserve", "reservation", "booking", "booked", "book the unit", "book a unit", "will book", "going to book",
        "deposit", "down payment", "eoi", "expression of interest", "cheque", "check ready", "sign the contract",
        "signing", "contracting", "chose the unit", "chose unit", "selected the unit", "unit selected",
        "ready to pay", "will pay", "paid", "hold the unit", "block the unit", "transfer the",
        "ha7gez", "hay7gez", "hya7gez", "hat7gez", "7agz", "7agaz", "7agazet", "ye7gez", "n7gez", "ne7gez",
        "3arboon", "3arboun", "3arbon", "mo2adam", "hayedfa3", "hyedfa3", "hadfa3", "dafa3", "shek", "ta3aqod", "ta3akod",
        "حجز", "هيحجز", "هتحجز", "يحجز", "تحجز", "نحجز", "عربون", "دفع المقدم", "هيدفع", "هتدفع", "دفع", "شيك",
        "تعاقد", "هيمضي", "يمضي العقد", "اختار الوحده", "اختار وحده", "اختارت الوحده", "اختار الشقه", "جاهز يشتري",
        "جاهز للحجز", "هيقفل", "قفلنا", "استمارة",
    ],
    "unresp": [
        "wait for respon*", "waiting for respon*", "waiting response", "waiting his reply", "waiting her reply",
        "no answer", "no asnwer", "no answe", "not answering", "didn't answer", "didnt answer", "doesn't answer",
        "doesnt answer", "not answer", "no reply", "not replying", "didn't reply", "didnt reply", "no response",
        "not responding", "unreachable", "not reachable", "switched off", "phone off", "closed phone", "phone closed",
        "can't reach", "cant reach", "couldn't reach", "couldnt reach", "ignored", "ignoring", "seen", "blocked",
        "out of coverage", "mbyrodsh", "mabyrodsh", "mbyrdsh", "mardsh", "maradesh", "mrdsh", "maraddish", "mrdish",
        "mafesh rad", "mafish rad", "msh byrod", "mesh byrod", "mabtrodsh", "mbtrodsh", "mardetsh", "mo3lak", "ma2fool",
        "maqfool", "mafool", "mesh beyrod", "msh by rod", "na,", "n/a",
        "مردش", "مردتش", "ماردش", "مبيردش", "مابيردش", "مبتردش", "مش بيرد", "مش بترد", "لا يرد", "لم يرد", "مفيش رد",
        "مقفول", "مغلق", "خارج التغطيه", "بلوك", "عمل سين", "سين بس", "محدش بيرد", "التليفون مش متاح", "التليفون مقفول", "مش بيرد",
        "مش بيرد على", "مبيرودش", "لا يجيب", "مش عارف اوصله", "مش عارفه اوصله",
    ],
    "hesit": [
        "yefakar", "byfkar", "byfakar", "bifakar", "byfker", "ramadan", "رمضان", "decision", "extend", "mekarar", "yeqarar",
        "ظروف", "zrof", "zoroof", "zorof", "السيوله", "سيوله", "cash flow", "liquidity", "seyola", "agl le", "2agl", "2agel", "akher elshahr", "akher elshar", "akher el shar",
        "moshkela", "moshkla", "problem", "مشكله",
        "cancel", "lesa msh", "lesa mesh", "lesa m4", "lesa byfakar", "lesa byfkar", "shahr gy", "shahr gay", "akher el shahr", "akhr el shahr", "awel el shahr", "اخر الشهر", "اول الشهر", "yesafer", "ysafer", "yelghi", "ابعتيلي واتساب", "ابعتلي واتساب", "يقرر", "hy2arar", "hay2arar", "yeqarar",
        "think", "will think", "consider", "hesitant", "hesitat*", "not sure", "expensive", "high price", "price is high",
        "prices are high", "over budget", "out of budget", "budget issue", "less budget", "discount", "location",
        "far", "consult", "postpone", "postponed", "later", "next month", "after eid",
        "travel", "abroad", "out of country", "outside egypt", "sick", "not ready", "not now", "compare", "comparing",
        "other options", "other project", "call back", "call me back", "busy", "will get back", "get back to",
        "hayfakar", "hatfakar", "hyfakar", "hatfkr", "hfakar", "yfakar", "hayfkr", "hyfkr", "hatfaker", "hayfaker",
        "ghaly", "ghali", "ghalya", "el se3r", "el s3r", "ba3id", "b3eed", "ba3eed", "mesafer", "msafer", "safer",
        "msafra", "mesafra", "hayrod 3alaya", "trod 3laya", "trod 3alaya", "yrod 3alya", "hayrod", "hatrod",
        "ba3den", "ba3deen", "b3den", "el shahr el gay", "msh delwa2ty", "mesh delwa2ty", "msh dlw2ty",
        "hayshawer", "hyshawer", "ta3ban", "3ayan", "mashghool", "mash8ool",
        "هيفكر", "يفكر", "هتفكر", "بيفكر", "بتفكر", "هيرد عليا", "هترد عليا", "هيرد علي", "هيرجعلي", "هترجعلي",
        "غالي", "غاليه", "السعر عالي", "الاسعار عاليه", "الاسعار غاليه", "الميزانيه", "ميزانيته", "خصم", "الموقع",
        "بعيد", "هيشاور", "هتشاور", "مسافر", "مسافره", "برا مصر", "بره مصر",
        "هيرجع", "هترجع", "تعبان", "تعبانه", "عيان", "عيانه", "مريض", "اجلها", "اجله", "اجلنا", "اتاجل", "تاجيل", "بعدين",
        "الشهر الجاي", "لسه مش", "مش جاهز", "مش دلوقتي", "مش دلوقت", "بيقارن", "بتقارن", "عروض تانيه",
        "كلمني بعد", "هيكلمني", "هتكلمني", "مشغول", "مشغوله", "مستني يبيع", "مستنيه تبيع", "محتاج يبيع",
    ],
    "wait": [
        "waiting for", "waitting",
        "mashro3", "mshro3", "el project", "new project",
        "اسكندرا", "مستني فتح", "مستنيه فتح",
        "luanch", "lunch alex", "lonch", "لونش", "مشروع اسكندري", "مستني مشروع", "مستنيه مشروع", "هتستني", "waiting unit", "waiting for our project", "alex", "alexandria", "اسكندري", "m3lomat", "when info", "lma ykon",
        "waiting release", "waiting for release", "waiting the release", "waiting for the release", "wait release",
        "waiting launch", "waiting for launch", "waiting for the launch", "waiting new launch", "new launch",
        "upcoming launch", "next phase", "new phase", "waiting price", "waiting for price", "waiting for the price",
        "waiting availability", "waiting for availability", "no availability", "not available", "sold out",
        "waiting for details", "waiting details", "waiting for project", "waiting the project", "waiting for alex",
        "release", "launch", "nazlen orayb", "nazleen orayb", "hanenzel", "hnnzel", "lama nnzel", "lama nenzel",
        "el tar7", "al tar7", "tar7", "mestany el", "mstny el", "mafish mesa7at", "mafesh mesa7at", "mafish 67",
        "msh available", "mesh available", "mesh mawgod", "msh mawgod", "mafish available",
        "الطرح", "طرح", "اطروحات", "تفاصيل الطرح", "مستني الطرح", "مستنيه الطرح", "منتظر الطرح", "لما ينزل",
        "لما ننزل", "هننزل", "نازلين", "مفيش مساحات", "مفيش متاح", "الاسعار الجديده",
        "مستني الاسعار", "مستني التفاصيل", "مستني المشروع", "مستني تفاصيل", "مستنيه تفاصيل", "المرحله الجديده",
    ],
    "engaged": [
        "tafaseel", "tafasel", "tfasel", "تفاصيل", "atklmna", "اتكلمنا",
        "nekhtar", "ne5tar", "nkhtar", "tegy", "tigi", "yegy", "hayegi", "office", "el office",
        "send details", "sent the details", "details on whatsapp", "طلب تفاصيل", "طلبت تفاصيل", "طلب", "هبعتلو", "هبعتلوا",
        "sa2al", "s2al", "s2l", "sa2alet", "s2lat", "سالني", "سالتني", "سالت", "mohtma", "mohtm", "ميعاد", "معاد", "تاكيد", "batab3", "batb3", "btab3", "bntab3", "bntna2esh", "bnetna2ash", "yeshtery", "yshtry", "btdrs", "bydrs",
        "interested", "intrested", "intersted", "interest in", "int in", "meeting", "met", "visit", "site visit",
        "came to", "attended", "zoom", "sent offer", "sent the offer", "sent details", "sent brochure",
        "sent price", "sent him", "sent her", "follow up", "following up", "likes", "liked", "wants", "want",
        "looking for", "asked about", "asking about", "asked for", "will come", "coming", "appointment",
        "good call", "positive", "hot", "serious", "potential",
        "mohtam", "mehtam", "mohtama", "3ayez", "3ayz", "3awz", "3awez", "3ayza", "3awza", "7ab", "3agbo", "3agbha",
        "gay", "gaya", "gayaly", "gayely", "gayly", "hayegy", "hyegy", "hygy", "hatigi", "hateegy", "zyara", "zeyara",
        "beysa2al", "bys2al", "bysa2al", "bydwar", "bydawar", "byedawar", "btdawar", "ba3atlo", "b3tlo", "ba3tlo",
        "hab3atlo", "habaatlo", "hab3tlo", "ba3atelha", "3amalna meeting", "et2abelna",
        "مهتم", "مهتمه", "عايز", "عاوز", "عايزه", "عاوزه", "حابب", "حابه", "عجبه", "عجبها", "جاي", "جايه", "جايلي",
        "هيجي", "هتيجي", "هيزور", "ميتنج", "ميتينج", "اجتماع", "زياره", "زار المشروع", "زارنا", "قابلته", "قابلتها", "قابلت",
        "بعتله", "بعتلها", "هبعتله", "هبعتلها", "ارسلت", "بيسال", "بتسال", "سال علي", "بيدور", "بتدور", "متابعه",
        "هتابع", "بتابع", "كويس", "ايجابي", "يشتري", "يشترى", "عرض سعر", "شاف", "شافت",
    ],
}


def norm_ar(s):
    s = re.sub(r"[ً-ْـ]", "", s)                       # tashkeel + tatweel
    s = re.sub(r"[أإآ]", "ا", s).replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    return s


def clean_note(s):
    s = re.sub(r"<(br|/p|/li|/div|/h\d)[^>]*>", "\n", s or "", flags=re.I)
    s = re.sub(r"<li[^>]*>", "- ", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = _html.unescape(s).replace("\xa0", " ")
    return re.sub(r"\n\s*\n+", "\n", s).strip()


_AR = re.compile(r"[؀-ۿ]")
def _compile(words):
    pats = []
    for w in words:
        w = norm_ar(w.lower())
        if _AR.search(w):
            pats.append(r"(?<![؀-ۿ])(?:[وفبل]|ال)?" + re.escape(w))
        elif w.endswith("*"):
            pats.append(r"(?<![a-z0-9])" + re.escape(w[:-1]))
        else:
            pats.append(r"(?<![a-z0-9])" + re.escape(w) + r"(?:s|es|d|ed|ing)?(?![a-z0-9])")
    return re.compile("|".join(pats)) if pats else None
_RX = None
def _rx():
    global _RX
    if _RX is None: _RX = [(k, _compile(STANCE_KEYWORDS.get(k, []))) for k, _l, _h in STANCES]
    return _RX


def classify_note(raw):
    """→ (stance_key, evidence_line, matched_word). Reads the LATEST update first:
    walks the note from the last line upwards; the first line that says something decides."""
    text = clean_note(raw)
    if not text: return NO_NOTES[0], "", ""
    lines = [l.strip(" -•\t") for l in text.split("\n") if l.strip(" -•\t")]
    for line in reversed(lines):
        low = norm_ar(line.lower())
        for k, rx in _rx():
            m = rx.search(low) if rx else None
            if m: return k, line, m.group(0).strip()
    return NO_NOTES[0], lines[-1] if lines else "", ""

# ==================================================================

OPP_FIELDS = ("Id, Name, StageName, Amount, CloseDate, CreatedDate, AgeInDays, LastActivityDate, "
              "LastStageChangeInDays, HasOverdueTask, NextStep, Rating__c, Project_Name_Text__c, "
              "Type_that_the_lead_is_interested_in__c, Number_of_Meetings__c, LeadSource, "
              "Account.Name, OwnerId, Owner.Name, Owner.IsActive, Owner.{tf}, " + ", ".join(NOTE_FIELDS))


def norm(s): return re.sub(r"\s+", " ", (s or "").strip()).lower()


# ============================ 1. FETCH ============================
def fetch(sf):
    """Returns (users, opps, open_stages, base_url) straight from Salesforce."""
    urows = sf.query_all(f"SELECT Id, Name, ManagerId, IsActive, {TEAM_FLD} FROM User")["records"]
    users = {r["Id"]: {"name": (r.get("Name") or "").strip(), "mgr": r.get("ManagerId"),
                       "active": r.get("IsActive", True),
                       "team": str(r.get(TEAM_FLD) or r.get(TEAM_FLD[0].upper() + TEAM_FLD[1:]) or "").strip()}
             for r in urows}
    stages = OPEN_STAGES
    if not stages:
        srows = sf.query_all("SELECT MasterLabel, ApiName, IsClosed, SortOrder FROM OpportunityStage "
                             "WHERE IsActive = true ORDER BY SortOrder")["records"]
        stages = [r.get("ApiName") or r["MasterLabel"] for r in srows if not r["IsClosed"]]
    reg = f" AND {REGION_FLD} = '{REGION_VAL}'" if APPLY_REGION_FILTER else ""
    opps = sf.query_all(f"SELECT {OPP_FIELDS.format(tf=TEAM_FLD)} FROM Opportunity "
                        f"WHERE IsClosed = false{reg}")["records"]
    base = SF_BASE_URL or ("https://" + sf.sf_instance if getattr(sf, "sf_instance", None) else "")
    return users, opps, stages, base


# ============================ 2. BUILD ============================
def note_hash(note):
    import hashlib
    return hashlib.sha1(note.strip().encode("utf-8")).hexdigest()[:12]


SHEET_COLS = ["Opportunity ID", "Stance", "Summary", "Note hash", "Reviewed on"]


def _sheet_rows_to_labels(rows):
    valid = {k for k, _l, _h in STANCES} | {NO_NOTES[0]}
    out = {}
    for r in rows:
        oid = str(r.get("Opportunity ID") or "").strip()
        st = str(r.get("Stance") or "").strip()
        hh = str(r.get("Note hash") or "").strip()
        if hh.startswith("h:"): hh = hh[2:]
        if not oid or st not in valid or not hh: continue
        out[oid] = {"stance": st, "summary": str(r.get("Summary") or "").strip(), "note_hash": hh,
                    "reviewed": str(r.get("Reviewed on") or "").strip()}
    return out


_GC = None
def sheets_client():
    """gspread client: service account (GOOGLE_SERVICE_ACCOUNT_JSON = key text, or a path in
    GOOGLE_APPLICATION_CREDENTIALS) → otherwise Colab Google sign-in. Raises if neither works."""
    global _GC
    if _GC is not None: return _GC
    import gspread
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    key = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if key or (path and os.path.exists(path)):
        from google.oauth2.service_account import Credentials
        info = json.loads(key) if key else json.load(open(path))
        _GC = gspread.authorize(Credentials.from_service_account_info(info, scopes=scopes))
        return _GC
    from google.colab import auth            # only inside Colab
    auth.authenticate_user()
    from google.auth import default
    creds, _ = default()
    _GC = gspread.authorize(creds)
    return _GC


def _sheet_id(url):
    m = re.search(r"/d/([A-Za-z0-9_-]+)", url or "")
    return m.group(1) if m else None


def load_sheet_labels(url=None, tab=None):
    """Downloads the reviewed stances Claude writes to Google Sheets every morning.
    1st try: public CSV link (sheet shared 'Anyone with the link can view').
    2nd try: Colab Google sign-in + gspread. Returns {} if neither works."""
    import csv, io, urllib.parse, urllib.request
    url, tab = url or LABELS_SHEET_URL, tab or LABELS_SHEET_TAB
    if not url: return {}
    m = re.search(r"/d/([A-Za-z0-9_-]+)", url)
    if not m:
        print(f"   ⚠️ LABELS_SHEET_URL doesn't look like a Google Sheet link: {url}"); return {}
    sid, rows, how = m.group(1), None, ""
    try:
        csv_url = (f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv"
                   f"&sheet={urllib.parse.quote(tab)}&headers=1")
        txt = urllib.request.urlopen(csv_url, timeout=60).read().decode("utf-8-sig")
        rows = list(csv.DictReader(io.StringIO(txt)))
        if not rows or "Opportunity ID" not in rows[0]: rows = None
        else: how = "public link"
    except Exception:
        rows = None
    if rows is None:
        try:
            rows = sheets_client().open_by_key(sid).worksheet(tab).get_all_records()
            how = "Google sign-in"
        except Exception as e:
            print(f"   ⚠️ couldn't read the '{tab}' tab of the labels sheet ({type(e).__name__}: {e}) "
                  f"→ using built-in / local labels only"); return {}
    labels = _sheet_rows_to_labels(rows)
    last = max((v["reviewed"] for v in labels.values() if v["reviewed"]), default="?")
    print(f"   ☁️ {len(labels):,} reviewed stances from Google Sheets ({how}) · latest review {last}")
    return labels


def load_labels(path=None, use_sheet=True):
    """Reviewed labels: built into this script  <  Google Sheet (LABELS_SHEET_URL)  <  NOTE_LABELS_FILE.
    A local label only overrides the sheet if it was reviewed on the same day or later."""
    labels = {k: {"stance": v[0], "summary": v[1], "note_hash": v[2]} for k, v in EMBEDDED_LABELS.items()}
    sheet = load_sheet_labels() if use_sheet else {}
    labels.update(sheet)
    path = path or NOTE_LABELS_FILE
    n_local = 0
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f: extra = json.load(f)
        for k, v in extra.items():
            if k not in sheet or str(v.get("reviewed", "")) >= sheet[k].get("reviewed", ""):
                labels[k] = v; n_local += 1
    print(f"   📘 {len(labels):,} reviewed labels ({len(EMBEDDED_LABELS):,} built in"
          + (f" + {len(sheet):,} from Google Sheets" if sheet else "")
          + (f" + {n_local:,} from {path}" if n_local else "") + ")")
    return labels


def write_all_notes(opps, rows, meta, url=None, tab=None):
    """Writes every open opportunity (full note text + stance + hierarchy) to the ALL_NOTES_TAB tab."""
    url, tab = url or LABELS_SHEET_URL, tab or ALL_NOTES_TAB
    sid = _sheet_id(url)
    if not (sid and tab): return
    try:
        sh = sheets_client().open_by_key(sid)
        try: ws = sh.worksheet(tab)
        except Exception: ws = sh.add_worksheet(title=tab, rows=len(rows) + 10, cols=16)
        full = {o["Id"]: clean_note("\n".join(str(o.get(x) or "") for x in NOTE_FIELDS if o.get(x))) for o in opps}
        SL = {k: l for k, l, _h in STANCES + [NO_NOTES]}
        HN = ["Healthy", "Watch", "At risk"]
        head = ["Opportunity ID", "Opportunity", "Account", "Sales Manager", "Team Leader", "Sales Rep", "Stage",
                "Amount", "Expected close", "Days in stage", "Client stance", "Health", "Reviewed",
                "Where the client stands", "Note (full)", "Refreshed"]
        order = {"lost": 0, "unresp": 1, "none": 2, "hesit": 3, "wait": 4, "engaged": 5, "book": 6}
        data = [head]
        for r in sorted(rows, key=lambda r: (r["sm"], r["tl"], r["rep"], order.get(r["sk"], 9), -r["amt"])):
            note = full.get(r["id"], "")
            if len(note) > SHEET_CELL_MAX: note = "…" + note[-SHEET_CELL_MAX:]
            data.append([r["id"], r["n"], r["a"], r["sm"], r["tl"], r["rep"], r["st"], r["amt"] or 0, r["cd"],
                         r["dis"], SL.get(r["sk"], r["sk"]), HN[r["h"]], "Yes" if r["rv"] else "Not yet",
                         r["sum"] or (r["ev"] if not r["rv"] else ""), note, meta["generated"]])
        ws.clear()
        if ws.row_count < len(data): ws.add_rows(len(data) - ws.row_count)
        for i in range(0, len(data), 400):
            ws.update(values=data[i:i + 400], range_name=f"A{i + 1}", value_input_option="RAW")
        try: ws.freeze(rows=1)
        except Exception: pass
        print(f"   🗂️ {len(data) - 1:,} open opportunities with full notes written to the '{tab}' tab")
    except Exception as e:
        print(f"   ⚠️ couldn't write the '{tab}' tab ({type(e).__name__}: {e}) — needs edit access "
              f"(Colab sign-in, or a service account that is an Editor on the sheet)")


def export_labels(opps, labels, out_path):
    """Writes the reviewed stance of every open opportunity whose note still matches its review
    and that is NOT already built into this script → CSV with SHEET_COLS
    (what Claude pastes into the Note_Labels tab every morning)."""
    import csv
    n = 0
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(SHEET_COLS)
        for o in sorted(opps, key=lambda o: o["Id"]):
            raw = "\n".join(str(o.get(x) or "") for x in NOTE_FIELDS if o.get(x))
            note = clean_note(raw)
            lab = labels.get(o["Id"])
            if not note.strip() or not lab or lab.get("note_hash") != note_hash(note): continue
            emb = EMBEDDED_LABELS.get(o["Id"])
            if emb and emb[0] == lab["stance"] and emb[2] == lab["note_hash"]: continue   # already built in
            w.writerow([o["Id"], lab["stance"], lab.get("summary", ""), "h:" + lab["note_hash"],
                        lab.get("reviewed") or "2026-11-02"]); n += 1
    print(f"   📤 {n:,} reviewed stances exported → {os.path.abspath(out_path)}")


def build(users, opps, stages, today, base_url="", log=print, labels=None, to_review=None):
    """users: id → {name, mgr, active, team}. opps: raw Opportunity records."""
    labels = load_labels() if labels is None else labels
    to_review = [] if to_review is None else to_review
    def find_user(nm, role):
        if str(nm).startswith("005") and nm in users: return nm
        hits = [u for u, v in users.items() if norm(v["name"]) == norm(nm) and v.get("active", True)]
        if not hits:
            log(f"   🚨 {role} '{nm}' not found / inactive"); return None
        if len(hits) > 1:      # same name twice → the one who manages Team Leaders
            tl_n = {norm(n) for n in TEAM_LEADER_NAMES}
            hits.sort(key=lambda h: -sum(1 for v in users.values() if v["mgr"] == h and norm(v["name"]) in tl_n))
        return hits[0]

    SM = {u: users[u]["name"] for u in (find_user(n, "Sales Manager") for n in SALES_MANAGER_NAMES) if u}
    TL = {u: users[u]["name"] for u in (find_user(n, "Team Leader") for n in TEAM_LEADER_NAMES) if u}
    sm_by_name = {norm(v): k for k, v in SM.items()}

    tl_sm = {}
    for t in TL:
        ov = next((v for k, v in TL_SALES_MANAGER_OVERRIDE.items() if norm(k) == norm(TL[t])), None)
        if ov and norm(ov) in sm_by_name:
            tl_sm[t] = SM[sm_by_name[norm(ov)]]; continue
        cur, seen = users[t]["mgr"], set()
        while cur and cur in users and cur not in seen and cur not in SM:
            seen.add(cur); cur = users[cur]["mgr"]
        if cur in SM:
            tl_sm[t] = SM[cur]
        else:
            m = users.get(users[t]["mgr"], {}).get("name") or "no manager"
            tl_sm[t] = f"Direct to {m}"
            log(f"   ⚠️ Team Leader '{TL[t]}' has no Sales Manager above them → shown as 'Direct to {m}'")

    def team_to_tl(v):
        v = norm(v)
        if not v: return None
        for t, nm in TL.items():
            a = norm(nm)
            if a == v or a.startswith(v) or v.startswith(a): return t
        return None

    def d(s): return date.fromisoformat(s[:10]) if s else None

    rows, other = [], {}
    for o in opps:
        own = o.get("OwnerId") or ""
        oname = ((o.get("Owner") or {}).get("Name") or users.get(own, {}).get("name") or "—").strip()
        oteam = (o.get("Owner") or {}).get(TEAM_FLD) or (o.get("Owner") or {}).get("Team_for_Tracking_membership__c") \
                or users.get(own, {}).get("team") or ""
        if own in SM:
            sm, tl, rep = SM[own], "Owned by Sales Manager", oname
        elif own in TL:
            sm, tl, rep = tl_sm[own], TL[own], f"{oname} (TL)"
        else:
            t = team_to_tl(oteam)
            if t: sm, tl, rep = tl_sm[t], TL[t], oname
            else:
                sm, tl, rep = "Other owners", "No team", oname
                other[oname] = other.get(oname, 0) + 1

        created = d(o.get("CreatedDate"))
        la = d(o.get("LastActivityDate"))
        ref = la or created
        dsa = max(0, (today - ref).days) if ref else None
        dis = o.get("LastStageChangeInDays")
        if dis is None: dis = o.get("AgeInDays") or 0
        cd = d(o.get("CloseDate"))

        raw = "\n".join(str(o.get(f) or "") for f in NOTE_FIELDS if o.get(f))
        note = clean_note(raw)
        lab = labels.get(o["Id"])
        if lab and lab.get("note_hash") == note_hash(note):
            sk, ev, kw, rv, summ = lab["stance"], "", "", True, lab.get("summary", "")
        elif not note.strip():
            sk, ev, kw, rv, summ = NO_NOTES[0], "", "", True, "No notes"
        else:
            sk, ev, kw = classify_note(raw); rv, summ = False, ""
            to_review.append({"Opportunity ID": o["Id"], "Name": o.get("Name") or "", "Owner": oname,
                              "Stage": o.get("StageName") or "", "Provisional stance": dict((k, l) for k, l, _h in STANCES + [NO_NOTES])[sk],
                              "Note": note})
        h = STANCE_HEALTH[sk]

        rows.append({
            "id": o["Id"], "n": o.get("Name") or "", "a": ((o.get("Account") or {}).get("Name") or ""),
            "sm": sm, "tl": tl, "rep": rep, "st": o.get("StageName") or "",
            "amt": o.get("Amount") or 0, "cd": cd.isoformat() if cd else "",
            "cr": created.isoformat() if created else "", "age": o.get("AgeInDays") or 0,
            "dis": dis, "la": la.isoformat() if la else "", "dsa": dsa, "nola": la is None,
            "ot": bool(o.get("HasOverdueTask")), "h": h, "sk": sk, "ev": ev[:300], "kw": kw, "rv": rv, "sum": summ,
            "note": note[-4000:], "cdp": bool(cd and cd < today),
            "rt": o.get("Rating__c") or "", "pj": o.get("Project_Name_Text__c") or "",
            "ty": o.get("Type_that_the_lead_is_interested_in__c") or "",
            "mt": o.get("Number_of_Meetings__c") or 0, "src": o.get("LeadSource") or "",
            "ns": o.get("NextStep") or "",
        })
    if other:
        log(f"   ℹ️ {sum(other.values())} opps owned by people outside the configured teams "
            f"→ 'Other owners': {dict(sorted(other.items(), key=lambda kv: -kv[1]))}")
    extra = sorted({r["st"] for r in rows} - set(stages))
    if extra:
        log(f"   ⚠️ open stage(s) not in the stage list, added at the end: {extra}")
    meta = {
        "asof": today.isoformat(), "generated": datetime.now(TZ).strftime("%Y-%m-%d %H:%M"),
        "region": REGION_VAL if APPLY_REGION_FILTER else "All regions",
        "stages": list(stages) + extra, "currency": CURRENCY, "base": base_url,
        "sms": sorted(set(SM.values())),
        "stances": [[k, l, h] for k, l, h in STANCES] + [list(NO_NOTES)],
        "reviewed": sum(1 for r in rows if r["rv"]), "unreviewed": sum(1 for r in rows if not r["rv"]),
    }
    return rows, meta


# ============================ 3. RENDER ============================
def render(rows, meta, fragment=False):
    """fragment=True → page body only (for publishing); False → full standalone .html file."""
    data = json.dumps(rows, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    m = json.dumps(meta, ensure_ascii=False).replace("</", "<\\/")
    body = TEMPLATE.replace("__DATA__", data).replace("__META__", m)
    if fragment: return body
    return ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
            '</head><body>' + body + '</body></html>')


TEMPLATE = r"""<title>Open Pipeline Health</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Layout: sticky filter bar → summary (KPIs, stage funnel, risk reasons) → team breakdown tree → opportunity list; drawer for drill-down */
:root{
  --bg:#f5f6f4; --panel:#ffffff; --ink:#18201f; --muted:#5d6a68; --line:#dfe4e2; --soft:#eef1ef;
  --accent:#0d6670; --accent-soft:#dcefee;
  --ok:#2f8a4e; --ok-soft:#dff1e4; --warn:#b7791f; --warn-soft:#f8ecd3; --bad:#c2412d; --bad-soft:#f8e0db;
  --sans:"Instrument Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
  --mono:"IBM Plex Mono",ui-monospace,Menlo,Consolas,monospace;
  --shadow:0 1px 2px rgba(20,30,30,.06),0 4px 16px rgba(20,30,30,.05);
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#111615; --panel:#1a2120; --ink:#e7ecea; --muted:#9aa8a5; --line:#2c3735; --soft:#222b29;
  --accent:#4fb3bb; --accent-soft:#1d3a3c;
  --ok:#5cc07c; --ok-soft:#1d3426; --warn:#e0a84a; --warn-soft:#3a2f19; --bad:#ec7a65; --bad-soft:#3d221d;
  --shadow:none; color-scheme:dark}}
:root[data-theme="dark"]{
  --bg:#111615; --panel:#1a2120; --ink:#e7ecea; --muted:#9aa8a5; --line:#2c3735; --soft:#222b29;
  --accent:#4fb3bb; --accent-soft:#1d3a3c;
  --ok:#5cc07c; --ok-soft:#1d3426; --warn:#e0a84a; --warn-soft:#3a2f19; --bad:#ec7a65; --bad-soft:#3d221d;
  --shadow:none; color-scheme:dark}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 var(--sans);
  padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));padding-block:0 48px}
body::before{content:"";position:fixed;inset:0 0 auto;height:env(safe-area-inset-top,0px);z-index:30;background:var(--bg)}
.wrap{max-width:1320px;margin:0 auto;display:flex;flex-direction:column;gap:18px}
.num,.mono{font-family:var(--mono);font-variant-numeric:tabular-nums}
header.top{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:12px;padding-top:24px}
h1{font-size:26px;line-height:1.15;margin:0;font-weight:700;letter-spacing:-.01em;text-wrap:balance}
.sub{color:var(--muted);font-size:13px;margin-top:4px}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--muted)}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:middle}
.d0{background:var(--ok)}.d1{background:var(--warn)}.d2{background:var(--bad)}

.filters{position:sticky;top:env(safe-area-inset-top,0px);z-index:20;background:var(--bg);padding:10px 0;border-bottom:1px solid var(--line);
  display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.filters label{display:flex;flex-direction:column;gap:2px;font-size:10.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);min-width:0}
select,input[type=search]{font:13px var(--sans);color:var(--ink);background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:6px 8px;max-width:190px}
input[type=search]{width:190px}
select:focus-visible,input:focus-visible,button:focus-visible,.clk:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
button{font:500 13px var(--sans);cursor:pointer;border-radius:6px;border:1px solid var(--line);background:var(--panel);color:var(--ink);padding:6px 10px}
button.primary{background:var(--accent);border-color:var(--accent);color:var(--panel)}
.chips{display:flex;flex-wrap:wrap;gap:6px;width:100%}
.chip{font-size:12px;background:var(--accent-soft);color:var(--ink);border:none;border-radius:999px;padding:3px 10px}
.chip b{margin-left:6px;color:var(--muted)}

.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow);padding:16px;min-width:0}
.panel h2{font-size:13px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin:0 0 12px;font-weight:600}
.panel h2 small{text-transform:none;letter-spacing:0;font-weight:400;margin-left:6px}

.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px;text-align:left;display:flex;flex-direction:column;gap:2px;box-shadow:var(--shadow)}
.kpi .l{font-size:11.5px;color:var(--muted)}
.kpi .v{font:500 24px/1.15 var(--mono);font-variant-numeric:tabular-nums}
.kpi .s{font-size:12px;color:var(--muted)}
button.kpi{cursor:pointer}
.kpi.k0{border-top:3px solid var(--ok)}.kpi.k1{border-top:3px solid var(--warn)}.kpi.k2{border-top:3px solid var(--bad)}
.kpi.on{outline:2px solid var(--accent)}

.grid2{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,1fr);gap:18px}
.bar{display:flex;height:12px;border-radius:3px;overflow:hidden;background:var(--soft);min-width:60px}
.bar span{display:block;height:100%}
.b0{background:var(--ok)}.b1{background:var(--warn)}.b2{background:var(--bad)}
.stage{display:grid;grid-template-columns:minmax(0,190px) 1fr 74px 96px;gap:12px;align-items:center;padding:9px 8px;border-radius:8px;cursor:pointer;border:1px solid transparent;width:100%;background:none;text-align:left}
.stage:hover{background:var(--soft)}
.stage.on{border-color:var(--accent);background:var(--accent-soft)}
.stage .nm{font-weight:600;font-size:13.5px}
.stage .nm small{display:block;font-weight:400;color:var(--muted);font-size:11.5px}
.stage .track{position:relative}
.stage .track .bar{height:18px}
.r{text-align:right}
.reason{display:grid;grid-template-columns:1fr 60px 52px;gap:10px;align-items:center;padding:8px;border-radius:8px;cursor:pointer;width:100%;background:none;border:1px solid transparent;text-align:left}
.reason:hover{background:var(--soft)}
.reason.on{border-color:var(--accent);background:var(--accent-soft)}
.reason .t{font-size:13.5px}
.reason .t small{display:block;color:var(--muted);font-size:11.5px}
.pillbar{height:6px;border-radius:3px;background:var(--soft);margin-top:4px;overflow:hidden}
.pillbar i{display:block;height:100%;background:var(--bad)}
.pillbar.w i{background:var(--warn)}

.tbl-wrap{overflow-x:auto;margin:0 -4px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600;text-align:right;padding:8px;border-bottom:1px solid var(--line);white-space:nowrap;position:relative}
th:first-child,td:first-child{text-align:left}
th.sort{cursor:pointer}
th.sort:hover{color:var(--ink)}
td{padding:7px 8px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}
tr.lvl0 td{font-weight:600}
tr.lvl1 td:first-child{padding-left:26px}
tr.lvl2 td:first-child{padding-left:48px;font-weight:400}
tr.lvl2 td{color:var(--ink)}
tr.tot td{font-weight:700;border-top:2px solid var(--line);background:var(--soft)}
.tog{border:none;background:none;padding:0 6px 0 0;color:var(--muted);font:600 13px var(--sans)}
.nm-btn{border:none;background:none;padding:0;font:inherit;color:inherit;cursor:pointer;text-align:left}
.nm-btn:hover{color:var(--accent);text-decoration:underline}
.clk{cursor:pointer;border-radius:4px}
.clk:hover{background:var(--accent-soft);color:var(--accent)}
td.z{color:var(--muted)}
.pct{display:inline-block;min-width:46px;padding:1px 6px;border-radius:4px;font-family:var(--mono);font-size:12px}
.p0{background:var(--ok-soft);color:var(--ok)}.p1{background:var(--warn-soft);color:var(--warn)}.p2{background:var(--bad-soft);color:var(--bad)}
.tag{display:inline-block;font-size:11px;padding:1px 7px;border-radius:999px;margin:1px 2px 1px 0;white-space:nowrap}
.t0{background:var(--ok-soft);color:var(--ok)}.t1{background:var(--warn-soft);color:var(--warn)}.t2{background:var(--bad-soft);color:var(--bad)}
.tg{background:var(--soft);color:var(--muted)}
.olist td{white-space:nowrap}
.olist td.wrap-t{white-space:normal;min-width:180px;max-width:300px}
.olist td.tags{white-space:normal;min-width:240px;max-width:340px;text-align:left}
.olist td.ev{white-space:normal;min-width:240px;max-width:380px;text-align:left;font-size:12.5px;color:var(--ink)}
.olist td.ev mark{background:var(--accent-soft);color:var(--accent);border-radius:3px;padding:0 2px}
.notebox{white-space:pre-wrap;unicode-bidi:plaintext;background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:12px 14px;font-size:13.5px;line-height:1.6;margin-top:6px;max-height:420px;overflow:auto}
.notebox .hit{background:var(--accent-soft);outline:1px solid var(--accent);border-radius:3px}
.sdot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:7px;vertical-align:middle}
.olist tbody tr{cursor:pointer}
.olist tbody tr:hover td{background:var(--soft)}
.more{display:flex;justify-content:center;gap:10px;align-items:center;margin-top:12px;color:var(--muted);font-size:12.5px}
.empty{padding:28px;text-align:center;color:var(--muted)}
a{color:var(--accent)}

.drawer-bg{position:fixed;inset:0;background:rgba(10,16,16,.35);z-index:40}
.drawer{position:fixed;top:0;right:0;bottom:0;width:min(920px,100%);background:var(--panel);z-index:41;border-left:1px solid var(--line);
  display:flex;flex-direction:column;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);padding-right:env(safe-area-inset-right,0px)}
.drawer header{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:16px 18px;border-bottom:1px solid var(--line)}
.drawer h3{margin:0;font-size:18px}
.drawer .body{overflow:auto;padding:14px 18px;flex:1}
.kv{display:grid;grid-template-columns:170px 1fr;gap:6px 14px;font-size:13.5px}
.kv dt{color:var(--muted)}
.kv dd{margin:0}
.note{font-size:12px;color:var(--muted)}
.foot{font-size:12px;color:var(--muted);line-height:1.6}
@media (max-width:1000px){.kpis{grid-template-columns:repeat(3,minmax(0,1fr))}.grid2{grid-template-columns:1fr}}
@media (max-width:760px){.filters{position:static}}
@media (max-width:560px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.stage{grid-template-columns:minmax(0,1fr) 60px 80px}.stage .track{grid-column:1/-1;order:5}
  select,input[type=search]{max-width:none;width:100%}.filters label{flex:1 1 45%}.kv{grid-template-columns:1fr}}
@media (prefers-reduced-motion:no-preference){.drawer{animation:slide .18s ease-out}@keyframes slide{from{transform:translateX(24px);opacity:.6}}}
</style>

<div class="wrap">
  <header class="top">
    <div>
      <h1>Open Pipeline Health</h1>
      <div class="sub" id="sub"></div>
    </div>
    <div class="legend">
      <span><i class="dot d0"></i>Healthy</span><span><i class="dot d1"></i>Watch</span><span><i class="dot d2"></i>At risk</span>
    </div>
  </header>

  <div class="filters" id="filters">
    <label>Sales manager<select id="f-sm"></select></label>
    <label>Team leader<select id="f-tl"></select></label>
    <label>Sales rep<select id="f-rep"></select></label>
    <label>Stage<select id="f-st"></select></label>
    <label>Client stance<select id="f-sk"></select></label>
    <label>Health<select id="f-h"><option value="">All</option><option value="0">Healthy</option><option value="1">Watch</option><option value="2">At risk</option></select></label>
    <label>Rating<select id="f-rt"></select></label>
    <label>Project<select id="f-pj"></select></label>
    <label>Search<input type="search" id="f-q" placeholder="Opportunity, account, rep"></label>
    <button id="reset" type="button">Clear filters</button>
    <div class="chips" id="chips"></div>
  </div>

  <section class="kpis" id="kpis"></section>

  <div class="grid2">
    <section class="panel">
      <h2>Pipeline by stage <small>click a stage to filter</small></h2>
      <div id="stages"></div>
    </section>
    <section class="panel">
      <h2>Where clients stand <small>from reading each opportunity's notes · click to filter</small></h2>
      <div id="reasons"></div>
    </section>
  </div>

  <section class="panel">
    <h2>Team breakdown <small>▸ expands a level · click a name to filter · click any number to see those opportunities</small></h2>
    <div class="tbl-wrap"><table id="tree"></table></div>
  </section>

  <section class="panel">
    <h2>Opportunities <small id="olist-n"></small></h2>
    <div class="tbl-wrap"><table class="olist" id="olist"></table></div>
    <div class="more" id="more"></div>
  </section>

  <p class="foot" id="foot"></p>
</div>
<div id="drawer-root"></div>

<script>
const DATA = __DATA__;
const META = __META__;
const STAGES = META.stages, CUR = META.currency;
const HN = ["Healthy","Watch","At risk"];
const STANCES = META.stances;                       // [key, label, health]
const SL = Object.fromEntries(STANCES.map(x=>[x[0],x[1]])), SH = Object.fromEntries(STANCES.map(x=>[x[0],x[2]]));
const SORDER = Object.fromEntries(STANCES.map((x,i)=>[x[0],i]));
const DORDER = ["book","engaged","wait","hesit","unresp","lost","none"];      // display order: healthy → at risk
const SDISP = [...STANCES].sort((a,b)=>(a[2]-b[2]) || ((DORDER.indexOf(a[0])+99)%99 - (DORDER.indexOf(b[0])+99)%99));
const STANCE_DESC = {book:"EOI, cheque, deposit, or unit chosen and paying", engaged:"Two-way contact: meeting, visit, reviewing an offer",
  wait:"Waiting for a release, launch, the Alex project or an unavailable unit", hesit:"Postponing, travelling, selling first, price or money concern, comparing",
  unresp:"Latest attempts unanswered, phone off", lost:"Not interested, bought elsewhere, not a buyer", none:"Empty note, or only a requirement"};
const F = {sm:"",tl:"",rep:"",st:"",sk:"",h:"",rt:"",pj:"",q:""};
const $ = id => document.getElementById(id);
const esc = s => String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmtN = n => n.toLocaleString("en-US");
function fmtM(v){ if(!v) return "0"; const a=Math.abs(v);
  if(a>=1e9) return (v/1e9).toFixed(a>=1e10?1:2)+"B"; if(a>=1e6) return (v/1e6).toFixed(a>=1e8?0:1)+"M";
  if(a>=1e3) return (v/1e3).toFixed(0)+"K"; return String(Math.round(v)); }
const fmtD = s => { if(!s) return "—"; const d=new Date(s+"T00:00:00"); return d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}); };
const pctCls = p => p>=50?2:(p>=25?1:0);

function match(r, skip){
  if(F.sm && skip!=="sm" && r.sm!==F.sm) return false;
  if(F.tl && skip!=="tl" && r.tl!==F.tl) return false;
  if(F.rep && skip!=="rep" && r.rep!==F.rep) return false;
  if(F.st && skip!=="st" && r.st!==F.st) return false;
  if(F.h!=="" && skip!=="h" && String(r.h)!==F.h) return false;
  if(F.rt && skip!=="rt" && (r.rt||"Not rated")!==F.rt) return false;
  if(F.pj && skip!=="pj" && (r.pj||"Not set")!==F.pj) return false;
  if(F.sk && skip!=="sk" && r.sk!==F.sk) return false;
  if(F.q && skip!=="q"){ const q=F.q.toLowerCase(); if(!(r.n+" "+r.a+" "+r.rep+" "+r.id+" "+r.note).toLowerCase().includes(q)) return false; }
  return true;
}
const rowsFor = skip => DATA.filter(r=>match(r,skip));

function agg(list){
  const a={n:list.length,v:0,h:[0,0,0],st:{},sk:{},dis:0};
  for(const r of list){ a.v+=r.amt; a.h[r.h]++; a.st[r.st]=(a.st[r.st]||0)+1; a.sk[r.sk]=(a.sk[r.sk]||0)+1; a.dis+=r.dis; }
  a.avgDis = a.n? a.dis/a.n : 0; a.riskPct = a.n? 100*a.h[2]/a.n : 0;
  return a;
}
function hbar(h,n){ if(!n) return '<div class="bar"></div>';
  return '<div class="bar" title="'+HN.map((x,i)=>x+": "+h[i]).join(" · ")+'">'+h.map((c,i)=>c?'<span class="b'+i+'" style="width:'+(100*c/n)+'%"></span>':"").join("")+'</div>'; }

/* ---------- filters ---------- */
function fillSelect(id,key,vals,allLabel){
  const el=$(id), cur=F[key];
  el.innerHTML='<option value="">'+allLabel+'</option>'+vals.map(v=>'<option'+(v===cur?' selected':'')+' value="'+esc(v)+'">'+esc(v)+'</option>').join("");
  if(cur && !vals.includes(cur)){ F[key]=""; el.value=""; }
}
function uniq(list,fn){ return [...new Set(list.map(fn))].sort((a,b)=>a.localeCompare(b)); }
function refreshFilters(){
  const smOrder = l => { const u=uniq(l,r=>r.sm); return u.sort((a,b)=>(META.sms.includes(b)-META.sms.includes(a))||a.localeCompare(b)); };
  fillSelect("f-sm","sm",smOrder(rowsFor("sm")),"All sales managers");
  fillSelect("f-tl","tl",uniq(rowsFor("tl"),r=>r.tl),"All team leaders");
  fillSelect("f-rep","rep",uniq(rowsFor("rep"),r=>r.rep),"All reps");
  const st=new Set(rowsFor("st").map(r=>r.st)); fillSelect("f-st","st",STAGES.filter(s=>st.has(s)),"All open stages");
  { const el=$("f-sk"), have=new Set(rowsFor("sk").map(r=>r.sk));
    el.innerHTML='<option value="">All stances</option>'+SDISP.filter(x=>have.has(x[0])).map(x=>'<option value="'+x[0]+'"'+(F.sk===x[0]?' selected':'')+'>'+esc(x[1])+'</option>').join("");
    if(F.sk && !have.has(F.sk)){F.sk=""; el.value="";} }
  fillSelect("f-rt","rt",uniq(rowsFor("rt"),r=>r.rt||"Not rated"),"All ratings");
  fillSelect("f-pj","pj",uniq(rowsFor("pj"),r=>r.pj||"Not set"),"All projects");
  $("f-h").value=F.h;
  const names={sm:"Manager",tl:"Team leader",rep:"Rep",st:"Stage",sk:"Client stance",h:"Health",rt:"Rating",pj:"Project",q:"Search"};
  $("chips").innerHTML = Object.keys(F).filter(k=>F[k]!=="").map(k=>{
    const val = k==="h"?HN[+F.h]:(k==="sk"?SL[F.sk]:F[k]);
    return '<button type="button" class="chip" data-k="'+k+'">'+names[k]+": "+esc(val)+'<b aria-label="remove">✕</b></button>'; }).join("");
}
function setF(k,v){ F[k]=(F[k]===v?"":v); if(k==="sm"){F.tl="";F.rep="";} if(k==="tl"){F.rep="";} render(); }

/* ---------- KPIs ---------- */
function renderKpis(list){
  const a=agg(list), n=a.n||1;
  const tile=(cls,l,v,s,h)=> (h!==undefined
    ? '<button type="button" class="kpi '+cls+(F.h===String(h)?' on':'')+'" data-h="'+h+'">'
    : '<div class="kpi '+cls+'">')+'<span class="l">'+l+'</span><span class="v">'+v+'</span><span class="s">'+s+'</span>'+(h!==undefined?'</button>':'</div>');
  $("kpis").innerHTML =
    tile("","Open opportunities",fmtN(a.n),agg(list.filter(r=>r.amt)).n+" with an amount") +
    tile("","Pipeline value",fmtM(a.v),CUR+" · avg "+fmtM(a.v/(list.filter(r=>r.amt).length||1))) +
    tile("k0","Healthy",fmtN(a.h[0]),(100*a.h[0]/n).toFixed(0)+"% · "+CUR+" "+fmtM(sumV(list,0)),0) +
    tile("k1","Watch",fmtN(a.h[1]),(100*a.h[1]/n).toFixed(0)+"% · "+CUR+" "+fmtM(sumV(list,1)),1) +
    tile("k2","At risk",fmtN(a.h[2]),(100*a.h[2]/n).toFixed(0)+"% · "+CUR+" "+fmtM(sumV(list,2)),2) +
    '<button type="button" class="kpi k0'+(F.sk==="book"?' on':'')+'" data-sk="book"><span class="l">Ready to book</span><span class="v">'+fmtN(a.sk.book||0)+'</span><span class="s">'+CUR+' '+fmtM(list.reduce((s,r)=>s+(r.sk==="book"?r.amt:0),0))+' per notes</span></button>';
}
const sumV=(l,h)=>l.reduce((s,r)=>s+(r.h===h?r.amt:0),0);
function median(a){ if(!a.length) return 0; a=[...a].sort((x,y)=>x-y); const m=a.length>>1; return a.length%2?a[m]:Math.round((a[m-1]+a[m])/2); }

/* ---------- stages ---------- */
function renderStages(){
  const list=rowsFor("st"), max=Math.max(1,...STAGES.map(s=>list.filter(r=>r.st===s).length));
  $("stages").innerHTML = STAGES.map(s=>{
    const l=list.filter(r=>r.st===s), a=agg(l);
    return '<button type="button" class="stage'+(F.st===s?' on':'')+'" data-st="'+esc(s)+'">'+
      '<span class="nm">'+esc(s.replace(/^Qualified\s*-?\s*/,""))+'<small>avg '+a.avgDis.toFixed(0)+' days in stage</small></span>'+
      '<span class="track"><span style="display:block;width:'+(100*a.n/max)+'%;min-width:'+(a.n?'6px':'0')+'">'+hbar(a.h,a.n)+'</span></span>'+
      '<span class="r num">'+fmtN(a.n)+'</span><span class="r num" title="'+CUR+'">'+fmtM(a.v)+'</span></button>';
  }).join("") + '<div class="note" style="margin-top:8px;padding-left:8px">Bars show count of opportunities, coloured by health. Right column = pipeline value ('+CUR+').</div>';
}

/* ---------- client stance (from notes) ---------- */
function renderReasons(){
  const list=rowsFor("sk"), n=list.length||1, col=["var(--ok)","var(--warn)","var(--bad)"];
  let last=-1, h="";
  for(const [k,l,hh] of SDISP){
    if(hh!==last){ h+='<div class="note" style="margin:'+(last<0?'0':'10px')+' 8px 2px;text-transform:uppercase;letter-spacing:.06em;font-size:10.5px">'+HN[hh]+'</div>'; last=hh; }
    const c=list.filter(r=>r.sk===k).length;
    h+='<button type="button" class="reason'+(F.sk===k?' on':'')+'" data-sk="'+k+'"><span class="t"><i class="sdot" style="background:'+col[hh]+'"></i>'+esc(l)+'<small style="padding-left:15px">'+STANCE_DESC[k]+'</small>'+
      '<div class="pillbar" style="margin-left:15px"><i style="width:'+(100*c/n)+'%;background:'+col[hh]+'"></i></div></span>'+
      '<span class="r num">'+fmtN(c)+'</span><span class="r num note">'+(100*c/n).toFixed(0)+'%</span></button>';
  }
  $("reasons").innerHTML = h;
}

/* ---------- tree ---------- */
const OPEN = new Set();
let TREE_SORT = {k:"n",d:-1};
const COLS = [
  {k:"n",l:"Open",f:a=>fmtN(a.n),v:a=>a.n,q:()=>true},
  {k:"v",l:"Value",f:a=>fmtM(a.v),v:a=>a.v,q:()=>true},
  ...STAGES.map(s=>({k:"st:"+s,l:s.replace(/^Qualified\s*-?\s*/,""),f:a=>a.st[s]?fmtN(a.st[s]):"–",v:a=>a.st[s]||0,q:r=>r.st===s})),
  {k:"hb",l:"Health mix",html:a=>hbar(a.h,a.n),v:a=>a.h[0]/(a.n||1)},
  {k:"h2",l:"At risk",f:a=>fmtN(a.h[2]),v:a=>a.h[2],q:r=>r.h===2},
  {k:"rp",l:"At risk %",html:a=>a.n?'<span class="pct p'+pctCls(a.riskPct)+'">'+a.riskPct.toFixed(0)+'%</span>':"",v:a=>a.riskPct,q:r=>r.h===2},
  ...SDISP.map(([k,l])=>({k:"sk:"+k,l:l,f:a=>a.sk[k]?fmtN(a.sk[k]):"–",v:a=>a.sk[k]||0,q:r=>r.sk===k})),
  {k:"dis",l:"Avg days in stage",f:a=>a.avgDis.toFixed(0),v:a=>a.avgDis},
];
function groupBy(list,key){ const m=new Map(); for(const r of list){ if(!m.has(r[key])) m.set(r[key],[]); m.get(r[key]).push(r);} return m; }
function sortGroups(m){ const c=COLS.find(c=>c.k===TREE_SORT.k)||COLS[0];
  return [...m.entries()].map(([k,l])=>[k,l,agg(l)]).sort((x,y)=>{
    const ox=x[0]==="Other owners"?1:0, oy=y[0]==="Other owners"?1:0; if(ox!==oy) return ox-oy;
    return TREE_SORT.d*(c.v(x[2])-c.v(y[2])) || x[0].localeCompare(y[0]); }); }
function cells(a,path){ return COLS.map((c,i)=>{
  if(c.html) return '<td>'+c.html(a)+'</td>';
  const val=c.v(a); if(!c.q || !val) return '<td class="'+(val?'num':'num z')+'">'+c.f(a)+'</td>';
  return '<td class="num clk" tabindex="0" role="button" data-path="'+esc(path)+'" data-col="'+i+'">'+c.f(a)+'</td>'; }).join(""); }
function renderTree(){
  const list=rowsFor(), head='<thead><tr><th>Sales manager / Team leader / Rep</th>'+COLS.map(c=>'<th class="'+(c.v?'sort':'')+'" data-k="'+esc(c.k)+'">'+esc(c.l)+(TREE_SORT.k===c.k?(TREE_SORT.d<0?' ↓':' ↑'):'')+'</th>').join("")+'</tr></thead>';
  let body="";
  for(const [sm,l1,a1] of sortGroups(groupBy(list,"sm"))){
    const p1=JSON.stringify([sm]), o1=OPEN.has(p1);
    body+='<tr class="lvl0"><td><button type="button" class="tog" data-open=\''+esc(p1)+'\' aria-expanded="'+o1+'">'+(o1?"▾":"▸")+'</button><button type="button" class="nm-btn" data-fk="sm" data-fv="'+esc(sm)+'">'+esc(sm)+'</button></td>'+cells(a1,p1)+'</tr>';
    if(!o1) continue;
    for(const [tl,l2,a2] of sortGroups(groupBy(l1,"tl"))){
      const p2=JSON.stringify([sm,tl]), o2=OPEN.has(p2);
      body+='<tr class="lvl1"><td><button type="button" class="tog" data-open=\''+esc(p2)+'\' aria-expanded="'+o2+'">'+(o2?"▾":"▸")+'</button><button type="button" class="nm-btn" data-fk="tl" data-fv="'+esc(tl)+'">'+esc(tl)+'</button> <span class="note">'+groupBy(l2,"rep").size+' reps</span></td>'+cells(a2,p2)+'</tr>';
      if(!o2) continue;
      for(const [rep,l3,a3] of sortGroups(groupBy(l2,"rep")))
        body+='<tr class="lvl2"><td><button type="button" class="nm-btn" data-fk="rep" data-fv="'+esc(rep)+'">'+esc(rep)+'</button></td>'+cells(a3,JSON.stringify([sm,tl,rep]))+'</tr>';
    }
  }
  body+='<tr class="tot"><td>Total</td>'+cells(agg(list),"[]")+'</tr>';
  $("tree").innerHTML=head+'<tbody>'+(list.length?body:'<tr><td colspan="'+(COLS.length+1)+'" class="empty">No opportunities match these filters.</td></tr>')+'</tbody>';
}
function pathRows(path){ const p=JSON.parse(path), keys=["sm","tl","rep"];
  return rowsFor().filter(r=>p.every((v,i)=>r[keys[i]]===v)); }

/* ---------- opportunity list ---------- */
let OSORT={k:"h",d:-1}, OLIM=50;
const OCOLS=[
  {k:"n",l:"Opportunity",v:r=>r.n.toLowerCase()},
  {k:"rep",l:"Rep",v:r=>r.rep},
  {k:"tl",l:"Team leader",v:r=>r.tl},
  {k:"st",l:"Stage",v:r=>STAGES.indexOf(r.st)},
  {k:"amt",l:"Amount",v:r=>r.amt},
  {k:"dis",l:"Days in stage",v:r=>r.dis},
  {k:"cd",l:"Expected close",v:r=>r.cd},
  {k:"h",l:"Client stance",v:r=>({lost:7,unresp:6,none:5,hesit:4,wait:3,engaged:2,book:1})[r.sk]||0},
  {k:"ev",l:"Where the client stands",v:r=>(r.sum||r.ev).toLowerCase()},
];
function sortList(list,S){ const c=OCOLS.find(c=>c.k===S.k)||OCOLS[0];
  return [...list].sort((a,b)=>{const x=c.v(a),y=c.v(b); return S.d*(x<y?-1:x>y?1:0) || b.amt-a.amt;}); }
function tags(r){ return '<span class="tag t'+r.h+'" title="'+HN[r.h]+'">'+esc(SL[r.sk])+'</span>'; }
function evHtml(r){ if(r.rv) return r.sum?'<span>'+esc(r.sum)+'</span>':'<span class="note">No notes</span>';
  if(!r.ev) return '<span class="note">No notes</span>';
  let t=esc(r.ev.length>220?r.ev.slice(0,220)+"…":r.ev);
  if(r.kw){ const i=r.ev.toLowerCase().indexOf(r.kw); if(i>=0 && i<220){ const k=esc(r.ev.slice(i,i+r.kw.length)); t=t.replace(k,'<mark>'+k+'</mark>'); } }
  return '<span class="tag tg" title="New or changed note, classified by word lists until it is reviewed">not reviewed</span> <span dir="auto">'+t+'</span>'; }
function oTable(list,S,lim,id){
  const head='<thead><tr>'+OCOLS.map(c=>'<th class="sort" data-ok="'+c.k+'" data-tbl="'+id+'">'+c.l+(S.k===c.k?(S.d<0?' ↓':' ↑'):'')+'</th>').join("")+'</tr></thead>';
  const rows=sortList(list,S).slice(0,lim).map(r=>'<tr data-id="'+r.id+'" tabindex="0">'+
    '<td class="wrap-t"><b>'+esc(r.n)+'</b>'+(r.a?'<div class="note">'+esc(r.a)+'</div>':'')+'</td>'+
    '<td>'+esc(r.rep)+'</td><td>'+esc(r.tl)+'</td><td>'+esc(r.st.replace(/^Qualified\s*-?\s*/,""))+'</td>'+
    '<td class="num">'+(r.amt?fmtM(r.amt):'<span class="note">—</span>')+'</td>'+
    '<td class="num">'+r.dis+'</td>'+
    '<td class="num">'+fmtD(r.cd)+'</td><td style="text-align:left">'+tags(r)+'</td><td class="ev">'+evHtml(r)+'</td></tr>').join("");
  return head+'<tbody>'+(rows||'<tr><td colspan="10" class="empty">No opportunities match these filters.</td></tr>')+'</tbody>';
}
function renderList(){
  const list=rowsFor();
  $("olist-n").textContent=fmtN(list.length)+" shown · most at risk first · click a row to read the full note";
  $("olist").innerHTML=oTable(list,OSORT,OLIM,"main");
  $("more").innerHTML = list.length>OLIM ? 'Showing '+fmtN(OLIM)+' of '+fmtN(list.length)+' <button type="button" id="showmore">Show 100 more</button><button type="button" id="showall">Show all</button>' : (list.length?'All '+fmtN(list.length)+' shown':'');
}

/* ---------- drawer ---------- */
let DR=null;
function closeDrawer(){ DR=null; $("drawer-root").innerHTML=""; document.body.style.overflow=""; }
function openList(title,list){ DR={type:"list",title,list,S:{k:"h",d:-1}}; drawDrawer(); }
function openOpp(id){ const r=DATA.find(x=>x.id===id); if(!r) return; DR={type:"opp",r,prev:DR&&DR.type==="list"?DR:null}; drawDrawer(); }
function drawDrawer(){
  document.body.style.overflow="hidden";
  let h='<div class="drawer-bg" id="dbg"></div><aside class="drawer" role="dialog" aria-modal="true" aria-label="Details"><header>';
  if(DR.type==="list"){
    const a=agg(DR.list);
    h+='<div><h3>'+esc(DR.title)+'</h3><div class="sub">'+fmtN(a.n)+' opportunities · '+CUR+' '+fmtM(a.v)+' · '+HN.map((x,i)=>a.h[i]+' '+x.toLowerCase()).join(' · ')+'</div></div>';
    h+='<button type="button" id="dclose">Close</button></header><div class="body"><div class="tbl-wrap"><table class="olist">'+oTable(DR.list,DR.S,5000,"drawer")+'</table></div></div>';
  } else {
    const r=DR.r, url=META.base?META.base.replace(/\/$/,"")+"/"+r.id:"";
    h+='<div><h3>'+esc(r.n)+'</h3><div class="sub">'+esc(r.a||"")+'</div><div style="margin-top:6px">'+tags(r)+'</div></div><div style="display:flex;gap:8px;flex-wrap:wrap">'+
      (DR.prev?'<button type="button" id="dback">← Back to list</button>':'')+'<button type="button" id="dclose">Close</button></div></header><div class="body">';
    h+='<p style="margin:0 0 4px"><b>Client stance: '+esc(SL[r.sk])+'</b> <span class="note">('+HN[r.h]+(r.rv?'':' · not reviewed yet, provisional')+')</span></p>';
    if(r.sum) h+='<p style="margin:0 0 10px;font-size:15px">'+esc(r.sum)+'</p>';
    if(r.note){ let nt=esc(r.note); if(r.ev && !r.rv){ const e=esc(r.ev); const i=nt.lastIndexOf(e); if(i>=0) nt=nt.slice(0,i)+'<span class="hit">'+e+'</span>'+nt.slice(i+e.length); }
      h+='<div class="notebox" dir="auto">'+nt+'</div><p class="note" style="margin:6px 0 16px">'+(r.rv?'Full note as written in Salesforce. Latest update is at the bottom.':'Highlighted line = the update the provisional stance was taken from.')+'</p>'; }
    else h+='<p class="note" style="margin:0 0 16px">This opportunity has no notes.</p>';
    h+='<dl class="kv">';
    const kv=[["Stage",r.st],["Amount",r.amt?CUR+" "+fmtN(r.amt):"Not set"],["Expected close date",fmtD(r.cd)+(r.cdp?" (passed)":"")],
      ["Days in current stage",r.dis],["Last activity",r.la?fmtD(r.la)+" ("+r.dsa+" days ago)":"Never — created "+r.dsa+" days ago"],
      ["Overdue task",r.ot?"Yes":"No"],["Age (days open)",r.age],["Created",fmtD(r.cr)],["Rating",r.rt||"Not rated"],["Project",r.pj||"Not set"],
      ["Unit type",r.ty||"—"],["Meetings logged",r.mt||0],["Lead source",r.src||"—"],["Next step",r.ns||"—"],
      ["Sales rep",r.rep],["Team leader",r.tl],["Sales manager",r.sm],["Opportunity ID",r.id]];
    h+=kv.map(([k,v])=>'<dt>'+k+'</dt><dd>'+esc(v)+'</dd>').join("")+'</dl>';
    if(url) h+='<p style="margin-top:16px"><a href="'+esc(url)+'" target="_blank" rel="noopener">Open in Salesforce ↗</a></p>';
    h+='</div>';
  }
  $("drawer-root").innerHTML=h+'</aside>';
  $("dclose").focus();
}

/* ---------- events ---------- */
document.addEventListener("click",e=>{
  const t=e.target.closest("button,td.clk,th.sort,tr[data-id],.chip");
  if(!t) return;
  if(t.id==="dbg"||t.id==="dclose") return closeDrawer();
  if(t.id==="dback"){ DR=DR.prev; return drawDrawer(); }
  if(t.id==="reset"){ for(const k in F) F[k]=""; $("f-q").value=""; return render(); }
  if(t.id==="showmore"){ OLIM+=100; return renderList(); }
  if(t.id==="showall"){ OLIM=1e9; return renderList(); }
  if(t.classList.contains("chip")){ const k=t.dataset.k; F[k]=""; if(k==="q") $("f-q").value=""; return render(); }
  if(t.dataset.h!==undefined) return setF("h",t.dataset.h);
  if(t.dataset.st!==undefined && t.classList.contains("stage")) return setF("st",t.dataset.st);
  if(t.dataset.sk) return setF("sk",t.dataset.sk);
  if(t.dataset.open){ const p=t.dataset.open; OPEN.has(p)?OPEN.delete(p):OPEN.add(p); return renderTree(); }
  if(t.dataset.fk){ const k=t.dataset.fk, v=t.dataset.fv; F[k]=(F[k]===v?"":v);
    if(k==="sm"){F.tl="";F.rep=""; if(F.sm) OPEN.add(JSON.stringify([v]));}
    if(k==="tl"){F.rep=""; const r=DATA.find(x=>x.tl===v); if(r&&F.tl){OPEN.add(JSON.stringify([r.sm])); OPEN.add(JSON.stringify([r.sm,v]));}}
    return render(); }
  if(t.matches("td.clk")){ const c=COLS[+t.dataset.col], p=JSON.parse(t.dataset.path);
    const list=pathRows(t.dataset.path).filter(c.q);
    return openList((p.length?p[p.length-1]:"All teams")+" — "+(c.k==="n"||c.k==="v"?"all open":c.l), list); }
  if(t.matches("th.sort")){
    if(t.dataset.k){ TREE_SORT = TREE_SORT.k===t.dataset.k?{k:t.dataset.k,d:-TREE_SORT.d}:{k:t.dataset.k,d:-1}; return renderTree(); }
    const S = t.dataset.tbl==="drawer"?DR.S:OSORT, k=t.dataset.ok;
    if(S.k===k) S.d=-S.d; else {S.k=k; S.d=(k==="n"||k==="rep"||k==="tl"||k==="cd"||k==="st")?1:-1;}
    return t.dataset.tbl==="drawer"?drawDrawer():renderList(); }
  if(t.matches("tr[data-id]")) return openOpp(t.dataset.id);
});
document.addEventListener("keydown",e=>{
  if(e.key==="Escape" && DR) return closeDrawer();
  if((e.key==="Enter"||e.key===" ") && e.target.matches("td.clk,tr[data-id]")){ e.preventDefault(); e.target.click(); }
});
for(const [id,k] of [["f-sm","sm"],["f-tl","tl"],["f-rep","rep"],["f-st","st"],["f-sk","sk"],["f-h","h"],["f-rt","rt"],["f-pj","pj"]])
  $(id).addEventListener("change",e=>{ F[k]=e.target.value; if(k==="sm"){F.tl="";F.rep="";} if(k==="tl") F.rep=""; render(); });
let qT; $("f-q").addEventListener("input",e=>{ clearTimeout(qT); qT=setTimeout(()=>{F.q=e.target.value.trim(); render();},180); });

function render(){ OLIM=50; refreshFilters(); const l=rowsFor(); renderKpis(l); renderStages(); renderReasons(); renderTree(); renderList(); }

$("sub").textContent = "As of "+fmtD(META.asof)+" · "+META.region+" · all open stages · generated "+META.generated+" Cairo";
$("foot").innerHTML = "<b>How health is decided.</b> Every opportunity's notes were read in full and the client placed in one stance with a one-line summary of where they stand. "+
  "<b>Healthy</b> = Ready to book or Engaged. <b>Watch</b> = Waiting on us or Hesitant. <b>At risk</b> = Unresponsive, Losing interest, or no clear update in the notes. "+
  fmtN(META.reviewed)+" notes reviewed"+(META.unreviewed?"; "+fmtN(META.unreviewed)+" new or changed notes have a provisional stance (marked “not reviewed”) until they are read":"")+". "+
  "Days in stage and expected close are shown for context only. Reps are grouped by their Team for Tracking membership; team leaders by their manager in Salesforce. Amounts in "+CUR+"; opportunities without an amount count as 0.";
render();
</script>
"""


# ===================== REVIEWED NOTE LABELS (2026-11-02) =====================
# Opportunity Id → [stance, one-line summary, note fingerprint]. Do not edit by hand;
# send Notes_To_Review_<date>.csv to Claude to have new / changed notes read.
EMBEDDED_LABELS = json.loads(r'''{"006PW00000oYsArYAK":["wait","Wants Alex launch; considering 1BR SouthMed too","eed8e97d589c"],"006PW00000sJUzNYAW":["wait","Wants Alexandria; waiting for project","e01307c950cc"],"006PW00000t07uhYAA":["unresp","Abroad, planned visit; latest call unanswered","dc649ec02498"],"006PW00000t0Je5YAE":["wait","Broker; follow up at releases","9ad4a3882af0"],"006PW00000t0MP3YAM":["wait","Met; waiting Privado prime-location release","3657f4bd2dff"],"006PW00000t0PzdYAE":["wait","Waiting for Alexandria release","735df8c94cb0"],"006PW00000t0mEiYAI":["wait","Waiting for Alex project details","0554f2033db6"],"006PW00000t0mPyYAI":["wait","Waiting for Alex","0c3239384bea"],"006PW00000t13x3YAA":["wait","Waiting for Alex project details","3ce4b2ddaf7f"],"006PW00000t1E7hYAE":["wait","Needs 2 apartments in Alexandria","b9581368f768"],"006PW00000t5EFaYAM":["none","No notes","da39a3ee5e6b"],"006PW00000t70lCYAQ":["wait","Waiting Alex release, 2BR","916926818cbd"],"006PW00000t7RlRYAU":["none","No notes","da39a3ee5e6b"],"006PW00000t7UbDYAU":["wait","Waiting Alex release","a904970c7bc8"],"006PW00000t7cthYAA":["wait","Waiting Alex release, 2BR","83aa9ec37f4a"],"006PW00000t81FNYAY":["wait","Waiting Alexandria release","f6ec881ede1b"],"006PW00000t9D3DYAU":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00000tI2enYAC":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00000tI7RaYAK":["wait","Bought SouthMed; call when Alex launches","e1b504ce20fd"],"006PW00000tIDGsYAO":["wait","Waiting for Alex project launch","e378bfefbc72"],"006PW00000tSzDlYAK":["wait","Interested in Alex only","0b7d0212409b"],"006PW00000trIDeYAM":["wait","Call her if Alex launches","52a99ec545a0"],"006PW00000tuqXLYAY":["wait","Waiting for Alex project","3ffdc0abb99b"],"006PW00000u4qcVYAQ":["wait","Wants 100m in Alex project","0e803dcdeb3b"],"006PW00000u56ZDYAY":["wait","Wants Alex project","1abd6484c514"],"006PW00000u5neEYAQ":["wait","Waiting Alex launch","e2b76921798e"],"006PW00000uDfYCYA0":["wait","Waiting for any Alex project","92f72c8a8b42"],"006PW00000uFSlKYAW":["wait","Interested in Alex project","ad340f07469a"],"006PW00000uFs4fYAC":["wait","Interested in Alex project","8fad6c6f5ab5"],"006PW00000uLtapYAC":["wait","Asked Alex details","a7108aa85cd6"],"006PW00000uLyAeYAK":["wait","Interested in Alex only","0b7d0212409b"],"006PW00000uM7x4YAC":["wait","Interested in Alex project","1abd6484c514"],"006PW00000uMBJ7YAO":["wait","Interested in Alex project","d33859e20795"],"006PW00000uMC5XYAW":["wait","Interested in Alex","970bae8791b8"],"006PW00000uMPQvYAO":["wait","Wants Alex; send details when out","da8d27ee9839"],"006PW00000uMPSXYA4":["wait","Interested in Alex project","1abd6484c514"],"006PW00000ufmQDYAY":["wait","Waiting for Alex details","8d7fc1a61baf"],"006PW00000ujliTYAQ":["wait","Waiting Alex release","a49f35cb9a42"],"006PW00000ujsIXYAY":["none","No notes","da39a3ee5e6b"],"006PW00000utio9YAA":["lost","Bought in Privado","0af23b4177a5"],"006PW00000v28VJYAY":["none","No notes","da39a3ee5e6b"],"006PW00000vCqLNYA0":["wait","Interested in Alexandria","b5c09d5a84bf"],"006PW00000vRBbFYAW":["hesit","Broker; still comparing resale and other projects","687380c27d8b"],"006PW00000viwcAYAQ":["none","No notes","da39a3ee5e6b"],"006PW00000vj8BtYAI":["wait","No answer; Alexandria interest","dcec92d1674c"],"006PW00000wj8L0YAI":["none","No notes","da39a3ee5e6b"],"006PW00000xD28UYAS":["wait","Waiting new Alex launch","129dc1308b54"],"006PW00000xD5r7YAC":["wait","Waiting Alex release","ae554d30ff9c"],"006PW00000xRPhuYAG":["wait","Interested in Alexandria","9b268fb7b6b2"],"006PW00000xaSyJYAU":["wait","Interested in Alexandria","1748e55ff076"],"006PW00000xhsb3YAA":["wait","Wants ready project in Alex; not answering","c3ef1c4358b9"],"006PW00000xi3OQYAY":["wait","Interested in Alexandria","dc17f54c7f9a"],"006PW00000xmE0HYAU":["wait","Needs Alex project","3502dc68966a"],"006PW00000yAnzOYAS":["wait","Interested in Alex project","1abd6484c514"],"006PW00000zLAVtYAO":["wait","Waiting Alex launch","1124b570266a"],"006PW00000zaASLYA2":["wait","Waiting for Alex details, 3-4BR","875251772755"],"006PW00000zrS6HYAU":["wait","Waiting Alex project release","ae516c7c7085"],"006PW000010RO5mYAG":["wait","Interested in Alex project","2bf2b33633d6"],"006PW000010VpH3YAK":["wait","Interested in Alex project","8fad6c6f5ab5"],"006PW00001196bOYAQ":["wait","Waiting Alex project launch","1e70b211d098"],"006PW000011vEekYAE":["none","No notes","da39a3ee5e6b"],"006PW000012YOxVYAW":["wait","Waiting Alex project","7bfc42462c2f"],"006PW0000138yBFYAY":["wait","Excited; waiting Alex prices and details","d937d2ef6fac"],"006PW000013KNO1YAO":["wait","Wants to buy in Alexandria","70c2c7a1494e"],"006PW000013YGxmYAG":["wait","Call when Alex project launches","10d91a62678b"],"006PW00001427tpYAA":["wait","Waiting B14 services-area units","6dd07a53cdca"],"006PW000014E00XYAS":["wait","Waiting for Alexandria","733a1043b355"],"006PW000014N2KrYAK":["wait","Wants something in Alexandria","d144498a2d98"],"006PW000014jxM9YAI":["wait","Asking about Alexandria","983c75c7665c"],"006PW000014v2ruYAA":["wait","Wants Alex project","23542297c87b"],"006PW000014vdjNYAQ":["none","No notes","da39a3ee5e6b"],"006PW0000155otfYAA":["wait","Wants Alex project","124131726ca9"],"006PW0000156henYAA":["none","No notes","da39a3ee5e6b"],"006PW000015E1QlYAK":["wait","Interested in Alex project","1abd6484c514"],"006PW000015zGVuYAM":["none","No notes","da39a3ee5e6b"],"006PW000016kFo5YAE":["wait","Call once Alex project exists","458443f3c318"],"006PW000016kyreYAA":["none","No notes","da39a3ee5e6b"],"006PW000016l0TdYAI":["engaged","Zoom booked; following twin house type D","03396701e850"],"006PW0000178fizYAA":["wait","Call when Alex has units","249dddca9dfe"],"006PW000017I9MrYAK":["wait","Call when Alex has units","610073826a27"],"006PW000017IP6fYAG":["wait","Waiting release; wants Alexandria","ae0e82ddc2be"],"006PW000017IdeHYAS":["wait","Needs apartment in Alexandria","7da7c8558ed8"],"006PW000017L3KvYAK":["engaged","Comparing Celia sizes; asked B14 200m details","92b1196da90c"],"006PW000017MpXGYA0":["wait","Phone closed; wants Alexandria","bdcc61ead4a6"],"006PW000017N2ivYAC":["wait","Wants Alex near Carrefour","25981603fe15"],"006PW000017N4JMYA0":["wait","Interested in Alex release, 3BR","a79c7a81d669"],"006PW000017N9dxYAC":["engaged","Wants Alex; open to Noor/Rehab offers","66ab951f5844"],"006PW000017NNnOYAW":["wait","Interested in Alexandria","5b60d1ba5bff"],"006PW000017NVHvYAO":["wait","Waiting Alex launch","aaefb4c92920"],"006PW000017NYyxYAG":["wait","Will wait for Alex project","f930f3acb7ea"],"006PW000017NgtGYAS":["wait","Asking about Alexandria","71c4b238a037"],"006PW000017Niv4YAC":["wait","Investment apartment in Alex when launched","6944066a084a"],"006PW000017NjPiYAK":["wait","Eagerly waiting Alex 2BR","fcd5e6dda7bc"],"006PW000017Nqm9YAC":["wait","Interested in Alexandria","a8ec17b632b6"],"006PW000017cNYqYAM":["wait","Waiting for 2BR in Alex","bc9eb7cf141a"],"006PW000017dDaLYAU":["wait","Waiting for Alex project","322260d8ebfb"],"006PW000017e2GIYAY":["hesit","Owns Sharm Bay; postponed to April (liquidity)","559a8f707b13"],"006PW000017eV3oYAE":["unresp","Wants Alex; not answering","9129d2700a26"],"006PW000017fBkXYAU":["wait","Call back when Alex details out","c0b48230a505"],"006PW000017pA42YAE":["wait","Waiting Alex release, 110m 3BR","ac5b8d00de0b"],"006PW0000180kIgYAI":["unresp","Abroad, asks Alex; number unreachable","4e6672e67ca2"],"006PW0000180zXqYAI":["wait","Waiting for Alexandria","15c4737b26f1"],"006PW0000181ETfYAM":["wait","Waiting Alex release","1edf281413d0"],"006PW0000182eQrYAI":["wait","Wants apartment in Alexandria","da36b6607d4e"],"006PW00001835PTYAY":["wait","Asking about Alexandria","24598e89de2a"],"006PW000018LVDNYA4":["wait","Waiting Alex release","f1c519773eb2"],"006PW000018LWUPYA4":["wait","Waiting Alex release","40d6a24509d8"],"006PW000018Lq3JYAS":["lost","Already bought; follow up for Alex","4594693fbe0e"],"006PW000018Q5TaYAK":["unresp","No answer","71b1b695cbd9"],"006PW000018UJL7YAO":["wait","Interested in Alexandria","fa0ecc8befef"],"006PW000018VkfmYAC":["hesit","Abroad; won't decide before year end","65e3cbcaa274"],"006PW000018c7EfYAI":["wait","Waiting Alex launch","f277c2df0cf9"],"006PW000018dORzYAM":["none","No notes","da39a3ee5e6b"],"006PW000018doshYAA":["hesit","Abroad, can't travel; collecting down payment","79c89adfb564"],"006PW000018e5K2YAI":["wait","Interested in Alex project","cd93a4f0d4ee"],"006PW000018mj90YAA":["unresp","Wants Zayed/Alex; not answering","09c3605d6061"],"006PW000018mtzZYAQ":["wait","Waiting for Alex project","b75f16db20f2"],"006PW000018ortdYAA":["wait","Changed mind; now wants Alexandria","2991f0b4bae2"],"006PW000019IimTYAS":["wait","Waiting for Alex project","3f5923bf334e"],"006PW000019McDuYAK":["lost","Location outside Alex not suitable","7f21f3336bb1"],"006PW000019PLRmYAO":["wait","Interested in Alexandria only","36e5988f1588"],"006PW000019UDBrYAO":["unresp","No answer, no WhatsApp","93479b643f89"],"006PW000019fNhKYAU":["wait","Waiting Alex project details","ccfef7164cae"],"006PW000019nH22YAE":["none","No notes","da39a3ee5e6b"],"006PW000019nW2fYAE":["wait","Interested in Alex project","124131726ca9"],"006PW000019nubFYAQ":["wait","Waiting Alex launch","530833c1492c"],"006PW00001AOXasYAH":["engaged","Zoom on Celia; also wants Alex","b0790696c7df"],"006PW00001AiGlyYAF":["wait","Call at Alex release","626da210c510"],"006PW00001AnzSIYAZ":["wait","Waiting Alex launch, 2-3BR","14e1a74f6cef"],"006PW00001AywalYAB":["unresp","Asked Alex/Noor/SouthMed; no reply","ae0bcf661539"],"006PW00001Az5sXYAR":["wait","Waiting Eden 70m","3dc0f2535ff9"],"006PW00001BJ98nYAD":["wait","SouthMed too pricey; waiting Alex","c2a82aaef7f9"],"006PW00001BObKaYAL":["wait","Waiting Alex release","6a393069d90c"],"006PW00001BQUZvYAP":["wait","Wants Alexandria","bc06a2283506"],"006PW00001BQaicYAD":["wait","Wants 2BR in Alex project","92dd6214ece7"],"006PW00001Btl7WYAR":["wait","Waiting for Alex project","fb5da86bd8b7"],"006PW00001BwXG8YAN":["wait","Wants Alexandria","259756d131c3"],"006PW00001BxBevYAF":["wait","Wants Alexandria not Cairo","9c4a3aa0aa13"],"006PW00001BxJKjYAN":["wait","Wants Alexandria","19140539474f"],"006PW00001CL6pNYAT":["none","No notes","da39a3ee5e6b"],"006PW00001CLQPuYAP":["hesit","Dislikes location; wants near club","93d1605f7a21"],"006PW00001CaCjBYAV":["wait","Owner; thinking investment, waiting Alex","9fa9de3e86ca"],"006PW00001E1307YAB":["hesit","Not ready until Dec/Jan; far distance","a819c04cbbbc"],"006PW00001EMEtOYAX":["hesit","Visit end of month; studying with husband","7ceeb755dc78"],"006PW00001Ea6khYAB":["wait","Interested in Alex, ~120m","285418942464"],"006PW00001EaThbYAF":["wait","Waiting Alex release details","c2f27b706e8b"],"006PW00001EaYCKYA3":["wait","Waiting for Alex project","5b6c8b5ef0fb"],"006PW00001Em73SYAR":["engaged","Existing client; may buy again","c1658a0bfecc"],"006PW00001EyKxmYAF":["wait","Will wait for Alex project","3154ee5e14fc"],"006PW00001FEaSJYA1":["wait","Waiting for Alex project","5b6c8b5ef0fb"],"006PW00001FIChVYAX":["engaged","Discussing Noor 119m details","cd99ada24270"],"006PW00001FIOp7YAH":["wait","Interested in Alex project","2bf2b33633d6"],"006PW00001FOJYDYA5":["wait","Interested in Alex, 3BR","f1d4d2c13c3d"],"006PW00001FVd8nYAD":["none","No notes","da39a3ee5e6b"],"006PW00001Fw8gnYAB":["wait","Waiting Alex launch","e729210841d7"],"006PW00001GB68VYAT":["unresp","Wants Alex; phone closed","db378a47c4d1"],"006PW00001GLbMFYA1":["wait","Waiting small gold-shop unit","6a5fcdf99c26"],"006PW00001GVvbOYAT":["unresp","Phone switched off","574a94601a1d"],"006PW00001GYpXpYAL":["wait","Still interested; waiting Alex details","0165641444f7"],"006PW00001H1IyIYAV":["book","Transferring 150K; awaiting Sharm Bay commercial","8cfa7585d16d"],"006PW00001H98X7YAJ":["engaged","Existing buyer; considering Spine office/1BR","156ec2a95667"],"006PW00001HIuZLYA1":["hesit","Shop in Spine; ~50% likely, depends on new prices","3a51c477964e"],"006PW00001HXwitYAD":["hesit","Unit sent (B14 garden); thinking with wife","17505fa510c3"],"006PW00001Hc1AVYAZ":["wait","Wants Alex; also interested Downtown project","a2e58e30824b"],"006PW00001Hi39ZYAR":["hesit","Abroad; meeting postponed until return","522105840441"],"006PW00001Hi6PBYAZ":["hesit","Busy/travelling; waiting for her return","fadbe333ee82"],"006PW00001Hxv1ZYAR":["wait","Interested in Alex project","01aafd51bcae"],"006PW00001I5h7vYAB":["wait","Discussed 3BR 162m; waiting Alex project","534ecfbf6023"],"006PW00001I9ceIYAR":["none","No notes","da39a3ee5e6b"],"006PW00001I9oynYAB":["wait","Wants 3BR in Alexandria","a2585f10acec"],"006PW00001IE7OXYA1":["book","Unit chosen; left a cheque but hesitant","3dbbbd9dc5e0"],"006PW00001IcrlBYAR":["hesit","Came to office, preparing down payment; father died, pause a week","7ef89b599af1"],"006PW00001J8JZuYAN":["wait","Interested in Alex project","92f32b4f439d"],"006PW00001JC4VqYAL":["wait","Will wait for Alex launch","530934b22d29"],"006PW00001JG389YAD":["wait","Wants Alex if price/plan suit","b00cced2f53d"],"006PW00001JXEdRYAX":["wait","Wants Alex, not near Carrefour","1c7fad168945"],"006PW00001JkuADYAZ":["wait","Waiting Alex launch","fc15b7328583"],"006PW00001K6JQYYA3":["none","No notes","da39a3ee5e6b"],"006PW00001K6pztYAB":["wait","Wants small shop; commercial only","b2676f243f91"],"006PW00001K6vu1YAB":["wait","Wants Alexandria","779b76fe94d1"],"006PW00001K7J0IYAV":["wait","Wants Alex project info","460100db2ce9"],"006PW00001KMFpDYAX":["wait","Interested in Alex","05a7ce221028"],"006PW00001KepZHYAZ":["wait","Wants Alexandria","fd9eb14fc8cf"],"006PW00001Kfa6mYAB":["hesit","Visited Spine with wife; travelling, will come on return","5b33fd7c5f47"],"006PW00001KrxAnYAJ":["hesit","Chose T1 twin house; in recovery after accident, still interested","81118e485c58"],"006PW00001L1xMQYAZ":["unresp","Waiting Alex; phone closed","672dde405b7c"],"006PW00001LAwSjYAL":["none","No notes","da39a3ee5e6b"],"006PW00001LCDEbYAP":["hesit","Chose V3 twin house; recovering from accident, still interested","f960f7d974b5"],"006PW00001LCMEfYAP":["hesit","Chose Spine office; recovering from accident, still interested","327e97fb4ee0"],"006PW00001LlGHeYAN":["wait","Low budget for SouthMed; wants Alexandria","872aa4d45b74"],"006PW00001M7HBrYAN":["wait","Waiting Alex project","1285e865061b"],"006PW00001NI9ofYAD":["engaged","Asked Lagoon chalet prices; zoom planned","3db333242033"],"006PW00001NIeKbYAL":["wait","Wants a shop in SouthMed","d7841fbcec5b"],"006PW00001NOG97YAH":["none","No notes","da39a3ee5e6b"],"006PW00001NOJmvYAH":["none","Only requirement: 3BR ground","e9e530527915"],"006PW00001NaiMXYAZ":["hesit","Asks cash discount occasionally; sporadic replies","0e3dccc6381f"],"006PW00001OLeNYYA1":["wait","Wants small 1BR; 66m too pricey; waits release","b1a3e0aa92d9"],"006PW00001OhyaTYAR":["hesit","Postponed","b444b460f149"],"006PW00001P3ct3YAB":["wait","Zoomed; deciding vs Alex purchase; waiting new Alex project","a7320b73ea2f"],"006PW00001PU3G5YAL":["engaged","Buying 2 twin houses; reviewing options sent","6b52c074d15f"],"006PW00001Q798jYAB":["hesit","Serious but abroad and busy; visit will take time","6c0e2c8a851b"],"006PW00001QZoLpYAL":["none","Only requirement: 32m","ab9de2256e53"],"006PW00001QkE3hYAF":["wait","Waiting Alex launch","248135083f2c"],"006PW00001R72gAYAR":["wait","Wants Alex; call when details out","c0b48230a505"],"006PW00001R8eLCYAZ":["hesit","Very hesitant; price high, not buying now","7f5d6fcdbb48"],"006PW00001RQUooYAH":["hesit","Visited, compared competitor; postponed for now","5b404f806529"],"006PW00001RQgWcYAL":["engaged","Back from Dubai; zoom for investment this weekend","1bc436bc75d4"],"006PW00001RSDYXYA5":["wait","Wants shop; required size not available","07420c665123"],"006PW00001S005qYAB":["hesit","Asked to call in December","b1e0f506b912"],"006PW00001S7dqRYAR":["engaged","Interested 1BR Lagoon; details sent","0d8e3765bab3"],"006PW00001S8l2tYAB":["wait","Left cheque; waiting new lake-view Privado release","0fe59eedba66"],"006PW00001SIHg5YAH":["engaged","Interested Privado 84m; reviewing payment","1329d452869f"],"006PW00001SImOvYAL":["lost","Bought already; not interested in Cairo","16bd3c72e7d2"],"006PW00001SP5NJYA1":["engaged","SouthMed only; resale 61m unit sent","a54597afcd27"],"006PW00001SgMptYAF":["wait","Wants 3BR ground; call at next release","f21cdc655c43"],"006PW00001SrH0UYAV":["wait","Will reserve as soon as Alex opens","7d6a5759625a"],"006PW00001T233pYAB":["wait","Still waiting Alex project","6717ea89d6c7"],"006PW00001T2Tb0YAF":["hesit","Postponed","0618fff1bf9f"],"006PW00001TRBB9YAP":["none","No notes","da39a3ee5e6b"],"006PW00001TUj9mYAD":["wait","Interested in Alex project","58f651677695"],"006PW00001Tj61dYAB":["wait","Waiting Privado 3BR release","4f17a4c14075"],"006PW00001U6GRhYAN":["wait","EOI refunded; waiting new Sahel release","db88144cad5b"],"006PW00001UL5ULYA1":["wait","Waiting large Noor unit or resale","fb01293be334"],"006PW00001UrDQLYA3":["engaged","Interested in clinic in Noor","2c5a9fcf196e"],"006PW00001V0v2LYAR":["engaged","Wants 3BR new type; meeting tomorrow","35135c4c1d87"],"006PW00001V1PEvYAN":["hesit","Meetings done; contract concern; travelling","976281e46f75"],"006PW00001V1QanYAF":["none","No notes","da39a3ee5e6b"],"006PW00001V1RQPYA3":["none","No notes","da39a3ee5e6b"],"006PW00001VKxCrYAL":["wait","Considered EOI; waiting studio release","277f2f7f4104"],"006PW00001VL1OXYA1":["book","Paid 100K EOI for garden villa","d6d11173066e"],"006PW00001VLQLKYA5":["wait","Waiting SouthMed studio release","5a57d22a88b6"],"006PW00001VPB4LYAX":["wait","Waiting Alex to reserve","1fd19636fc0c"],"006PW00001VPZbJYAX":["hesit","Liked project; postponed, transfer issue at work","e9f9609c4c6a"],"006PW00001VQUHqYAP":["engaged","Reviewing SouthMed 2BR ground 90m","ca170e13ae34"],"006PW00001VQYdIYAX":["engaged","Wants ground 3BR; needs plan and payment first","ac7d68cbf3ac"],"006PW00001VezyVYAR":["wait","Interested in new ground chalet release","6414d563e7a9"],"006PW00001VlVsPYAV":["engaged","Busy now; promised to close tomorrow on Privado","1becea398f3b"],"006PW00001VqqhhYAB":["book","Paid 100K EOI for garden villa","e74defae16da"],"006PW00001W1QokYAF":["wait","Very interested in Alex project","45dc16245323"],"006PW00001W2f2XYAR":["wait","Sold apartment; ready to book Alex","6854bf5be60f"],"006PW00001W5Ie6YAF":["engaged","Existing client's brother; pitching Noor/Madinaty","400475445e1c"],"006PW00001W5LnFYAV":["hesit","Thinking Privado vs Noor; no liquidity now","0e2ea6af8e2f"],"006PW00001WdbmLYAR":["engaged","Arranging meeting at Alex branch","db0e56d6505b"],"006PW00001WtcEbYAJ":["hesit","Met; comparing alternatives, undecided","61667bed8e12"],"006PW00001Wu2LxYAJ":["unresp","Wants Alex; phone unreachable","7dc39adc8cef"],"006PW00001X4jZJYAZ":["wait","Wants 3BR ground chalet; new type","a2ebb5aa31ea"],"006PW00001X8NK5YAN":["unresp","Budget mismatch; not answering","292b6241eb7d"],"006PW00001XIlGHYA1":["wait","Waiting next release, 2BR ground","d64d4f29d3ba"],"006PW00001XdxO6YAJ":["hesit","Prices too high","4c745c1f9cd6"],"006PW00001XiF4gYAF":["unresp","Wants Alex; no answer","1dc5ced38d66"],"006PW00001XiX81YAF":["wait","Wants Alex project","2423d83f8222"],"006PW00001XibrZYAR":["wait","Two months away; may consider Alex","bfdc752e28a0"],"006PW00001Xp6RGYAZ":["wait","Very interested in Alex project","505c8dd8ade5"],"006PW00001YAckfYAD":["engaged","Interested Lagoon/Privado; may visit Mon/Tue","47afcab0e6d4"],"006PW00001YFH2rYAH":["wait","Waiting for Alex project","322260d8ebfb"],"006PW00001YGJ7RYAX":["unresp","Thinking after meeting; now not answering","aec002bc160a"],"006PW00001ZDRx1YAH":["wait","Wants ~150m in Alex project","f0fdc47fe8df"],"006PW00001ZDnxWYAT":["engaged","Zoom done on 2BR Lagoon; reply in 2 days","e4b0acde6a35"],"006PW00001ZE0Q6YAL":["wait","Negotiating villas; waiting new senior chalet release","3bffbc16dbd2"],"006PW00001ZEAu6YAH":["wait","Bought via another branch; now waiting Alex","52f82a02edf6"],"006PW00001ZEhJlYAL":["hesit","Selling her villa first; still interested SouthMed","1aaace2fb880"],"006PW00001ZEtEXYA1":["wait","Waiting Alex project release","02b6a9df1d91"],"006PW00001ZJcUjYAL":["hesit","Buying with sister; circumstances, meeting pending","efe254dae7e3"],"006PW00001ZJpOjYAL":["wait","Repeat buyer; ready when Alex info out","fc955929d197"],"006PW00001ZJqPdYAL":["hesit","Waiting to sell Aswan property first","9e085c022611"],"006PW00001ZJrtZYAT":["hesit","Money issue abroad; meeting postponed, may buy Nov","cf3a598e64e1"],"006PW00001ZK9k1YAD":["hesit","Asked 2 more weeks; worried about market","8c1a19c65e6e"],"006PW00001ZLVYqYAP":["unresp","Wants B14 2BR; latest follow-up unanswered","56411cd1e5e4"],"006PW00001ZWTOnYAP":["hesit","Coming from Canada in Jan; wants 15-yr plan","1d88d13fab8f"],"006PW00001ZkWBVYA3":["wait","Waiting Alexandria project","1c61d59f9379"],"006PW00001ZkaOpYAJ":["unresp","Comparing plans with family; no response since","a5ec19528e3d"],"006PW00001ZkkcgYAB":["unresp","Not answering at all","643c24ec295a"],"006PW00001ZkpKbYAJ":["engaged","Liked Celia townhouse; thinking, will reply","704a2177b0c3"],"006PW00001Zl7MLYAZ":["wait","Wants Alexandria apartment","94047fb96351"],"006PW00001ZykxlYAB":["wait","Twin house in Alex project","b91fd5a4e67e"],"006PW00001ZzAnNYAV":["wait","Waiting Alex release details","71ae3c35c092"],"006PW00001a7G4QYAU":["none","Only requirement: 1BR Marina","6af7509db1c6"],"006PW00001a9MwwYAE":["hesit","Family illness; keeps postponing, still in touch","750a7ec4a219"],"006PW00001aAtSbYAK":["engaged","Discussed 2BR Lagoon prices; offer sent","810a390aaa63"],"006PW00001aBYNlYAO":["unresp","Latest follow-up unanswered","b08d809b668a"],"006PW00001aBhIyYAK":["none","Only requirement: studios","612e6e576ce9"],"006PW00001aRqq6YAC":["engaged","Meeting done on 2BR; following up","a330d55a6e52"],"006PW00001aSEafYAG":["none","Only requirement: T6 first-row villa","e65045311e44"],"006PW00001aSRPpYAO":["hesit","Liked Noor; switched to SouthMed; back in Feb","91b776f9bb69"],"006PW00001aScbPYAS":["wait","Waiting small open-air shop","96c32bb3f98f"],"006PW00001aTqqnYAC":["engaged","Met; plans site visit and second meeting","0910017b389d"],"006PW00001aX0s1YAC":["wait","Waiting 3BR ground release","7f3cfbb09a95"],"006PW00001aXaXKYA0":["wait","Wants Alex launch","7a07c70e7a62"],"006PW00001aXkbVYAS":["wait","Waiting front-facing Sharm Bay 1BR release","bdbdcc896877"],"006PW00001aYAR9YAO":["hesit","Travelling; call Monday night","7fcb663a89cd"],"006PW00001aajsbYAA":["wait","Waiting Alex project release","c999e2179c09"],"006PW00001abBAXYA2":["wait","Still interested, waiting Alex","06a20809b6fc"],"006PW00001abSphYAE":["none","No notes","da39a3ee5e6b"],"006PW00001abYgbYAE":["none","Not registered / no info","905b0ac7dcba"],"006PW00001ababxYAA":["hesit","In Turkey; call back in a month","9f66ef22ac72"],"006PW00001abeZ8YAI":["engaged","Two zooms done; liked Spine; offers sent","1d5bc9b2c87c"],"006PW00001acjGcYAI":["wait","Waiting senior chalet release","0f7353c8d925"],"006PW00001ahnppYAA":["hesit","Busy; asked to call Wednesday","29537fc43dde"],"006PW00001airD4YAI":["wait","Abroad; waiting C4 release","c4c88bfe4342"],"006PW00001am3j9YAA":["engaged","Discussed project details; following up","5502c61cf049"],"006PW00001amXm1YAE":["wait","Waiting Alex project","87d21d69f6f1"],"006PW00001ayYRNYA2":["wait","Waiting Alex project","f424fe46d49b"],"006PW00001azzCcYAI":["unresp","Wants Madinaty 3BR; phone off, no reply","9666c0fa7648"],"006PW00001bA5dLYAS":["lost","Owner checking prices only; not answering","d7a49bb96be7"],"006PW00001bZdBzYAK":["engaged","Likes Marina; wants offers for comparison","e02b2b74e849"],"006PW00001baCHlYAM":["engaged","Iraqi Madinaty owner; meet in October","7a68033edcb8"],"006PW00001baD2XYAU":["wait","Bought in Mountain View; wants Alex details","0de9de8fc4f8"],"006PW00001baXsTYAU":["engaged","Zoom done on 2BR; reply tomorrow","aba1f39fe437"],"006PW00001baeM5YAI":["wait","Will invest in Alex land","3346d9e5b24d"],"006PW00001bakL3YAI":["wait","Waiting Alex project","22f8443791e2"],"006PW00001basdZYAQ":["wait","Still waiting Alex","e43ee70697c2"],"006PW00001bayhNYAQ":["engaged","Investor; zoom scheduled","3defabb3cd8e"],"006PW00001beNXxYAM":["engaged","Wants Madinaty RTM typical floor","eebae22ff7d9"],"006PW00001bpdMTYAY":["engaged","Following up; unit sent","50f68b9efd2a"],"006PW00001bpedVYAQ":["wait","Will wait for Alex project","1953734886e0"],"006PW00001bvfa7YAA":["wait","Will wait for Lagoon garden release","371d7c6e21ad"],"006PW00001c0UHRYA2":["none","No notes","da39a3ee5e6b"],"006PW00001c0rdpYAA":["wait","Interested in Alex project","e7870d602bd1"],"006PW00001cBbVjYAK":["wait","Waiting Alex launch","63442c16a069"],"006PW00001cC1WcYAK":["wait","Will wait for Alexandria","3dd41f15aa95"],"006PW00001cCtrxYAC":["engaged","Zoom planned on Lagoon/Marina","6c8791a332d9"],"006PW00001cF2zBYAS":["hesit","Decided but can't pay until selling","cb99c2c7e0a2"],"006PW00001cW2aOYAS":["wait","Zoom set; will buy in Alex project","c78c10539f71"],"006PW00001cW4WjYAK":["hesit","Will act in December on Sharm","d7c3f7c4ee36"],"006PW00001cW5cUYAS":["hesit","Death in family; asked two weeks","35e48dcf7488"],"006PW00001cWiXNYA0":["none","Only requirement: 60-70m retail Noor","61abe5d2938b"],"006PW00001cX1RyYAK":["hesit","Comparing residential vs commercial; in Dubai","842cea35c992"],"006PW00001cbNYkYAM":["wait","Interested 1BR at Alex launch","70e282bf44b7"],"006PW00001cfeazYAA":["unresp","Postponed to year end; now not answering","6b9f536699d9"],"006PW00001cfvySYAQ":["engaged","Meeting planned end of month","dbcf88ed279d"],"006PW00001cgZXeYAM":["hesit","Negotiated Madinaty types; next month","26fe925abeb8"],"006PW00001clwesYAA":["hesit","Consulting family; deposit release issue","b31987719069"],"006PW00001cmGSHYA2":["wait","Waiting commercial, 18M budget","df74bd8ee9e6"],"006PW00001cmaQzYAI":["engaged","Met on Noor 82m; thinking","e3131ebbc670"],"006PW00001cphCcYAI":["unresp","Waiting Alex; latest call unanswered","c2f65d37ef35"],"006PW00001cwp6zYAA":["wait","Wants 61m Sharm in next release","493d9567fac0"],"006PW00001cwsjBYAQ":["wait","Wants small Spine commercial unit","4280c85abfdc"],"006PW00001cxQw9YAE":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00001cxjAoYAI":["engaged","Zoom done; considering 1BR Golf","bd9fe2328680"],"006PW00001cxsijYAA":["none","Broker; only requirement 1-2BR Lagoon","a4a5d2b7525f"],"006PW00001cyFQnYAM":["none","Only requirement: 2BR SouthMed","f18c7c2826d1"],"006PW00001d8AHlYAM":["wait","Not Madinaty; wants ready unit in Alex","189929a2ff74"],"006PW00001d8TAkYAM":["wait","Waiting Alexandria launch","c868c4095f97"],"006PW00001d9FP4YAM":["hesit","Still arranging budget; maybe next month","e6f1a5871182"],"006PW00001dXdOcYAK":["engaged","Noor 119m details; meeting this week","1b398b0ad44c"],"006PW00001dXhnBYAS":["lost","Chose unit then travelled; not interested","a711463c0918"],"006PW00001dXpEUYA0":["engaged","Following on Privado 2BR","4f465325d24f"],"006PW00001dXslrYAC":["wait","Can't join this release; wants next one","9de259507ed4"],"006PW00001dXuSgYAK":["wait","Contact when Alex project launches","aeb667a4aed3"],"006PW00001dY2A6YAK":["hesit","Interested but funds tied elsewhere; not now","d89105c9655d"],"006PW00001dY3kTYAS":["none","No notes","da39a3ee5e6b"],"006PW00001dcBEnYAM":["hesit","Waiting for his vacation to visit","7a943da093b4"],"006PW00001dciQsYAI":["hesit","Travelling a month; then to reserve","f21e2b4972d4"],"006PW00001ddf9CYAQ":["wait","Very interested; waiting Alex project and payment details","35b0cdf0da31"],"006PW00001deXmHYAU":["wait","From Alex; waiting Alex project","26370ffe93a5"],"006PW00001dev2EYAQ":["hesit","Must sell something first","be50d932918f"],"006PW00001dhUblYAE":["hesit","Bank transfer problem; trying to solve","1f6291db6026"],"006PW00001dmfLFYAY":["unresp","Waiting wife's return; latest call unanswered","0389bc82641d"],"006PW00001dxQvZYAU":["hesit","Postponed; call end of November","bd0e8e07d767"],"006PW00001dyx9XYAQ":["book","Buying 184m Noor with sister; cheques ready","858e0ad08528"],"006PW00001e3xQcYAI":["engaged","Repeat client; discussing Village and Marina","7e784b53be96"],"006PW00001e84GbYAI":["lost","Husband: not interested in Cairo; call for Alex","179efd64e745"],"006PW00001e8ID7YAM":["hesit","Interested in Noor; will close in November","d0098e05a101"],"006PW00001e8IOPYA2":["wait","Wants 3BR in Alexandria","50d701e00822"],"006PW00001e8WSzYAM":["unresp","Waiting Marina chalet; not answering","2157f94820c9"],"006PW00001e8pCHYAY":["hesit","No action before December","65f076f35f6d"],"006PW00001e9U7OYAU":["wait","Waiting Alex release details","7e38b5b06819"],"006PW00001eTEuVYAW":["wait","Wants ready unit in Madinaty; waiting new launch","7d0fec570b76"],"006PW00001eTkkfYAC":["book","Unit selected; solving bank account for cheques","17dfda3c612d"],"006PW00001eYwQDYA0":["unresp","Waiting Madinaty RTM release; no answer","9fae2341e46d"],"006PW00001enN7BYAU":["engaged","Interested 2-3BR; details sent, reply tomorrow","f227dab6a634"],"006PW00001exksCYAQ":["hesit","Met; needs to sell Aswan land first","b215386cb0a0"],"006PW00001exmfPYAQ":["hesit","Will wait next release; she is ill","c57ce0122fac"],"006PW00001ey69SYAQ":["wait","Wants Noor shop (food)","e6d1d40bcf42"],"006PW00001f05UFYAY":["hesit","Met; busy, arranging affairs","0cf5b5e6e1ec"],"006PW00001f8V06YAE":["none","No notes","da39a3ee5e6b"],"006PW00001f8ubBYAQ":["wait","Chosen unit taken; waiting next Noor release","b13611082711"],"006PW00001fAeFOYA0":["engaged","Visited; comparing townhouse vs Privado","7b50b3f0bc1e"],"006PW00001fKdgDYAS":["book","Will transfer on 146m Privado at Thursday release","4a3cb4e105b2"],"006PW00001fSuNJYA0":["engaged","Meeting next Monday, 2BR","a635ec8e6636"],"006PW00001fSuaIYAS":["wait","Waiting Noor RTM 3BR release","cb5e9157126b"],"006PW00001fSyU9YAK":["hesit","Consulting his children; postponing","5fb5f7979f3e"],"006PW00001fTClfYAG":["engaged","Looking for Sharm Bay studio","3bb321a91623"],"006PW00001fU1oBYAS":["wait","Wants 3BR in Alex project","2c852d297ad0"],"006PW00001fU4ywYAC":["wait","Wants Alexandria apartment","dd89b0e072b8"],"006PW00001fUOxdYAG":["hesit","Chose unit; father in hospital, postponed to Dec","9b9cf23d80b0"],"006PW00001fYCBVYA4":["wait","Wants to buy in Alexandria","efc59f610c7a"],"006PW00001fYV7iYAG":["engaged","Meeting Saturday; Noor RTM interest","f22f2866436b"],"006PW00001fYbmbYAC":["wait","Will wait for Alex project","dc96b3feb9e2"],"006PW00001fYghRYAS":["wait","Waiting Alex release","b15922694315"],"006PW00001fYzP7YAK":["book","Choosing unit at office when she's free","b3abb3c10e11"],"006PW00001fZE9fYAG":["wait","Waiting Alex project","0e3e4f63e2be"],"006PW00001fZO49YAG":["none","Broker; only requirement 2BR garden Lagoon","e018e2421570"],"006PW00001fZd9fYAC":["none","No notes","da39a3ee5e6b"],"006PW00001fbmNKYAY":["engaged","Zoom done, wife visited site; deciding","2c96bff734b6"],"006PW00001fcHnhYAE":["hesit","Wants 2 units 131m; travelling, back 14 Nov","e35264ec9b3e"],"006PW00001fdwaJYAQ":["hesit","Postponed till 2 Nov","cbf0a3ecafb8"],"006PW00001foYMZYA2":["hesit","Not ready until November (funds)","289aeb2db97f"],"006PW00001fopAVYAY":["engaged","Met at Dokki; discussing Noor/SouthMed","117c6a61f671"],"006PW00001foyiPYAQ":["hesit","Comparing competitor; still calculating","3c1a00804254"],"006PW00001fp917YAA":["lost","Not interested; asked 2BR SouthMed earlier","b93b14cec893"],"006PW00001fq1fpYAA":["hesit","Down payment not ready; follow up in a month","a56522182a93"],"006PW00001fqQKrYAM":["wait","Won't transfer until release date known","fad428922fef"],"006PW00001fqjQjYAI":["hesit","Visited; won't buy before a month, travelling","64a3f5e5268e"],"006PW00001fqkWTYAY":["unresp","Site visit postponed; not answering","d68fbecc70d0"],"006PW00001frAXPYA2":["hesit","Interested but circumstances; always busy","a6a04694648f"],"006PW00001ftyzYYAQ":["unresp","Missed agreed call; sending Madinaty/Noor","44f41490f436"],"006PW00001fuXXRYA2":["none","Only requirement: 2BR Spine/Privado","69c9f2e153f0"],"006PW00001fvfBCYAY":["engaged","Recent SouthMed buyer; coming next week to view","0b5a1565330d"],"006PW00001fw7AfYAI":["engaged","Wants Noor shop; meeting when in Egypt","17354f854f49"],"006PW00001fw8zZYAQ":["engaged","Phone often off; may come this weekend","de4c9fd7d4cf"],"006PW00001fwALRYA2":["engaged","Sent 52m office option","a37bfbea0978"],"006PW00001fyhV8YAI":["hesit","EOI submitted; wife in hospital, waiting next release","4861e5f139bc"],"006PW00001fz0W1YAI":["hesit","Ready only in December","65a5531b592a"],"006PW00001fzEHFYA2":["wait","Will wait for Alex","82b071999211"],"006PW00001fzGpJYAU":["hesit","Spine investment; memory issue, sons to decide","f1dfcccee8ca"],"006PW00001fzTWPYA2":["hesit","In Sharm; postponed to end of month","9d6eefa1f7e5"],"006PW00001fzXRxYAM":["hesit","Selling Mansoura flat first","fb7a97f62219"],"006PW00001fzzuTYAQ":["hesit","In Dubai a month; then choose unit","cfdb093afeb0"],"006PW00001g0LdGYAU":["hesit","Postponed to start of year","d266bb100f5b"],"006PW00001g0ta6YAA":["hesit","Sick; meeting postponed 2 weeks","5e5da63601f5"],"006PW00001g0vooYAA":["unresp","Not answering; wanted Sharm studio","329678aede95"],"006PW00001g13cfYAA":["unresp","Disappeared; not answering us or broker","3d7a503036b0"],"006PW00001g1sldYAA":["hesit","Friend must sell land first","3b08b0dc04a7"],"006PW00001g1yENYAY":["hesit","Another month","f14f7883040c"],"006PW00001g24V4YAI":["hesit","Abroad 3 months; undecided","c8a5406a36d5"],"006PW00001g26bkYAA":["unresp","No answer","9b0c28c7f876"],"006PW00001g2shxYAA":["unresp","No answer","136415afc8d1"],"006PW00001g3FrSYAU":["hesit","Work circumstances; postponed to Nov/Dec","282c398893b0"],"006PW00001g5PmfYAE":["hesit","Selling London property first","4f168e3295b9"],"006PW00001gAB6MYAW":["none","Only requirement: Noor shop","ac3daccb0b20"],"006PW00001gANCMYA4":["engaged","Discussing typical floor (no ground left)","9c5d34c7d66b"],"006PW00001gAbooYAC":["hesit","Wants ground 3BR RTM; not ready","33c728de7c00"],"006PW00001gAhNxYAK":["unresp","Flight attendant; hard to reach","c9b98f7643b7"],"006PW00001gAj9eYAC":["engaged","Wants Noor townhouse with siblings","b357c8dc7e32"],"006PW00001gArDeYAK":["hesit","Abroad; undecided","c77b939842fd"],"006PW00001gAyrpYAC":["hesit","Postponed 3 months","a28f0943c142"],"006PW00001gB5LVYA0":["engaged","Following up; call in two weeks","2c2ae558d4e2"],"006PW00001gBhFRYA0":["wait","Wants sea-view 2BR; waiting new release","970bb15b3130"],"006PW00001gBtYLYA0":["wait","Waiting Alex project","2ae02184a929"],"006PW00001gByo9YAC":["wait","Abroad; waiting Alex details","e7132081ba32"],"006PW00001gC2mwYAC":["hesit","Conference 15 days; meeting next week","8dd3db9c0968"],"006PW00001gC76gYAC":["engaged","Asked 98m/119m prices; sending","63758c152239"],"006PW00001gCdWNYA0":["engaged","Asked Noor 82/98m; follow-up in 3 days","7b77b312b75c"],"006PW00001gD9CsYAK":["hesit","Still calculating; will call when in Cairo","ce47a86d71a4"],"006PW00001gE1RlYAK":["engaged","Will come within a month","1d4679aeda2a"],"006PW00001gVAgkYAG":["engaged","Business owner; follow up 28 Nov","522751b983c0"],"006PW00001gVSvNYAW":["unresp","Funds pending, in Alex; not answering","90f34b10beec"],"006PW00001gW430YAC":["wait","Wants RTM in Noor when it opens","dc83c5e3dab7"],"006PW00001gWiwUYAS":["unresp","Meeting held; now not answering","6286085648d5"],"006PW00001gZtovYAC":["unresp","Was very engaged; stopped responding suddenly","7416aa6bf6ec"],"006PW00001gb6x0YAA":["engaged","Noor 2BR; site and office visit","36b2c1c2053a"],"006PW00001gb8vZYAQ":["unresp","Liked quattro; waiting funds; not answering","2623452078d0"],"006PW00001gbAyzYAE":["lost","Bought elsewhere; not buying now","9fe24ba7b868"],"006PW00001gbDK9YAM":["hesit","Follow up in December","d96bf134436a"],"006PW00001gbDTpYAM":["wait","Investment 1-2BR in Alex","0b92844cd1c0"],"006PW00001gbDf7YAE":["hesit","Townhouse too pricey; money mid-November","afbf67e6818a"],"006PW00001gbIoUYAU":["wait","Interested in Alex","f00cb2f3fbba"],"006PW00001gbLUdYAM":["engaged","Wants Sharm shop to rent; coming next week","0fe904efa5a8"],"006PW00001gbLw2YAE":["hesit","Not ready to close now","3f004e61c4c5"],"006PW00001gbPErYAM":["hesit","Keeps postponing meetings","8d691f223738"],"006PW00001gbRYQYA2":["lost","Only gathering prices; won't buy","9871a6cf6cea"],"006PW00001gbWmcYAE":["unresp","Thinking after meeting; not answering","c3435ef77096"],"006PW00001gcXS6YAM":["wait","Annual-payment concern; joining new release","70f41c030e79"],"006PW00001gfNUlYAM":["engaged","Spine 122m; meeting this week","1390d241900f"],"006PW00001gfXk5YAE":["unresp","Waiting brother; repeated no answer","b15cfd048b00"],"006PW00001gfYbJYAU":["engaged","Quattro over budget; sending 3BR","415481a62d0c"],"006PW00001gfZh6YAE":["hesit","Abroad; back in 3 weeks","fe86dba39d68"],"006PW00001gfkKjYAI":["none","Only requirement: 141m B14","ce7e11e6e1f2"],"006PW00001gfn46YAA":["hesit","Thinking; waiting funds","64bb7eb8b5d2"],"006PW00001gfp97YAA":["hesit","Abroad; meet on return","d00c565a5130"],"006PW00001gfzQDYAY":["engaged","Liked cash price; preparing down payment","c54b4384e30f"],"006PW00001gg1n0YAA":["hesit","Waiting to sell his house","36d5a353cd30"],"006PW00001gg6MrYAI":["wait","Interested in Alex project","ad340f07469a"],"006PW00001gg7E5YAI":["hesit","Must sell Hyde Park flat first","8f80ad77cbc5"],"006PW00001ggANGYA2":["none","No notes","da39a3ee5e6b"],"006PW00001ggpJxYAI":["hesit","Must sell her flat first; no buyer yet","d7925882029e"],"006PW00001ghma9YAA":["engaged","Noor 3BR; following up","d07f9b561383"],"006PW00001gigKbYAI":["engaged","Busy; arranging meeting","813fee9b1096"],"006PW00001gknjOYAQ":["none","Only requirement: 3BR Noor 119","ca35518c588d"],"006PW00001glundYAA":["none","Only requirement: 2BR Noor","3165ad476028"],"006PW00001gp8kPYAQ":["hesit","Comparing projects; waiting until month start","4864ce7a4cf1"],"006PW00001gryjhYAA":["wait","Will transfer when Alex details out","3ccf66d6891c"],"006PW00001gsA05YAE":["wait","Will wait for Alex project","0f5c467a7f94"],"006PW00001gsGIPYA2":["hesit","Coming in a month","fbe236321082"],"006PW00001gsJBRYA2":["engaged","Reviewing twin houses with family","9202d70c710c"],"006PW00001gsiWRYAY":["wait","Selected 119m; now wants Noor studios release","1e84c405730b"],"006PW00001gtMs1YAE":["wait","Waiting for Alex project","7e7d465dc4e1"],"006PW00001gtQ7dYAE":["unresp","Asked 119m; not answering","27a343d6a451"],"006PW00001gtqYMYAY":["wait","Wants Alexandria","637a27b7adaa"],"006PW00001guKY1YAM":["unresp","Meeting missed; work pressure; no answer","9cd8d5689cb2"],"006PW00001guXqDYAU":["engaged","Existing Noor buyer; call at each launch","6f0174617f8f"],"006PW00001guptZYAQ":["wait","Prefers waiting for Alex project","16a6e3bdb8fc"],"006PW00001gvwkvYAA":["wait","Waiting for Noor 67m","5256235310dc"],"006PW00001gwIohYAE":["engaged","Wants two 146m Privado units; coming to pay DP cash","5ebbc2200d49"],"006PW00001gwUt3YAE":["engaged","Friend visited site; 1BR interest","35826da8f356"],"006PW00001gztjeYAA":["engaged","Wants RTM 123m Noor; details sent","ce90a000f9fc"],"006PW00001h2nOMYAY":["engaged","In Egypt 3 weeks; family bereavement","04168606b1aa"],"006PW00001h4bMHYAY":["unresp","Doctor, always busy; not answering","2332e305d51e"],"006PW00001h4ctRYAQ":["engaged","Following 1BR garden in Noor","70080aa61122"],"006PW00001h4e2PYAQ":["hesit","Abroad until 15 Nov","b876d70f546f"],"006PW00001h4nP3YAI":["engaged","Closing pending; 2BR ground La Cantine","4ed384cd647d"],"006PW00001h4oGHYAY":["wait","Won't buy now; waiting Alex project","3ab8dd5ab75b"],"006PW00001h50vlYAA":["hesit","Money tied in business; waiting end of month","60c65cff826d"],"006PW00001h52fpYAA":["lost","Personal circumstances; dropped the idea","8a435ef73556"],"006PW00001h55QnYAI":["unresp","Wants Privado view 3BR; not answering","dd9300ddcb87"],"006PW00001h5RKrYAM":["wait","Will wait for Alex project","5f66b3962cdf"],"006PW00001h5aXrYAI":["unresp","Visited Privado; stopped answering","457b3bcac837"],"006PW00001h5uY9YAI":["wait","Wants 3BR in Alex near Moharram Bek","4bf40863ad7b"],"006PW00001h5xXdYAI":["wait","Wants nursery; none available yet","c63e1a2b46c9"],"006PW00001h5zMXYAY":["hesit","Disappears and returns; still thinking","9379394a0d72"],"006PW00001h6BigYAE":["hesit","Travelling to Saudi; decide end of month with wife","4aa36cdacaad"],"006PW00001h6CDJYA2":["engaged","Coming in 2-3 weeks to view options","2641bb51d6d3"],"006PW00001h6GX5YAM":["lost","Not interested","d509d7c8a4c9"],"006PW00001h6UuzYAE":["hesit","Liquidity issues at work; postponed","994489490b71"],"006PW00001h6VpRYAU":["engaged","SouthMed 1BR ground; meeting Wednesday","535e19e554c6"],"006PW00001h77t9YAA":["engaged","Large apartment; Noor site visit planned","f48a8fe4e7e4"],"006PW00001hApnyYAC":["engaged","Site visit Sunday, 131m","d3319678b11c"],"006PW00001hAwfpYAC":["engaged","Meeting 21 Oct, 3BR Noor","592c38f0cb24"],"006PW00001hBDQYYA4":["hesit","Travelling; back end of week","d68f5758777a"],"006PW00001hGYfZYAW":["unresp","Planned visit; now not answering","472343d54438"],"006PW00001hHMm2YAG":["hesit","Interested; can't act for a month","e1fddc992581"],"006PW00001hHRSLYA4":["hesit","Contract payment issue; postponed next month","b3627cd9adbf"],"006PW00001hHTnVYAW":["hesit","Postponed to end of year","d6af8c03e9e5"],"006PW00001hI93ZYAS":["hesit","Met at Noor; ill, still thinking","706d5efdf022"],"006PW00001hJqHlYAK":["wait","Will wait for new Noor launch","ff33a3eb82f8"],"006PW00001hKPdhYAG":["engaged","Settled on Madinaty 116m; finalizing next week","1f137812c4db"],"006PW00001hKpt9YAC":["engaged","Meeting Sunday, Noor","ece0ffc7f46e"],"006PW00001hMJ9BYAW":["none","Only requirement: 2-3BR RTM; in Dubai","a2e69c5fc85f"],"006PW00001hdJEPYA2":["none","No notes","da39a3ee5e6b"],"006PW00001hdMk9YAE":["hesit","Not buying soon","34ad7435db52"],"006PW00001hdShVYAU":["wait","Will wait for new projects","26a4302d12fe"],"006PW00001hdZXiYAM":["engaged","Interested in Noor; call back","5c22d70c9ee1"],"006PW00001hdlyfYAA":["hesit","Met; brother calculating installments","08599890bdff"],"006PW00001hdqQTYAY":["hesit","Commitments; will act in January","4d5073401efe"],"006PW00001he1deYAA":["engaged","Met; coming again next week","2500b88a3aa9"],"006PW00001he4QFYAY":["hesit","Son abroad; no action until he visits","d5e80b690b2d"],"006PW00001he8jxYAA":["hesit","Transfer limit; coming to SouthMed next month","399142dfbc8c"],"006PW00001heKjVYAU":["unresp","Son ill; latest call unanswered","ab9a716f69d4"],"006PW00001heVDWYA2":["engaged","Considering 127m layout","af60b8681226"],"006PW00001heVyHYAU":["hesit","Liked unit; selling asset first, 2 weeks","14611655a543"],"006PW00001hedvqYAA":["hesit","No action until visiting Nov/Dec","12e2da996c2c"],"006PW00001hhHxCYAU":["hesit","Zoom done; decides end of month","d9815e157de3"],"006PW00001hilQCYAY":["hesit","Back in December","b269eace8013"],"006PW00001hiyooYAA":["wait","Asking about Alex project","7ddb3c54c8fe"],"006PW00001hj2sPYAQ":["engaged","Meeting Saturday","89033f075ff4"],"006PW00001hj46DYAQ":["engaged","Multi-unit investor; following","b21507efd6f8"],"006PW00001hj4PZYAY":["hesit","Meeting postponed a week","7353a9777bf6"],"006PW00001hj4qzYAA":["hesit","Needs time after Alamein purchase","0b98165f70d2"],"006PW00001hj55VYAQ":["wait","Repeat client; waiting Alex project","b932ee0954e5"],"006PW00001hj5gbYAA":["none","Only requirement: 2BR Noor 97m","3fea4b5dadae"],"006PW00001hj6pZYAQ":["hesit","Will buy at year end","8560279da4ae"],"006PW00001hjA5BYAU":["hesit","Not before February","dc6690e921b0"],"006PW00001hjBPRYA2":["hesit","Very low budget","f66620f2676a"],"006PW00001hjEwpYAE":["none","No notes","da39a3ee5e6b"],"006PW00001hjS0TYAU":["engaged","Meeting today, 3BR ground","17c9bee19311"],"006PW00001hjSq7YAE":["hesit","Nothing until next month","bfac29d3a890"],"006PW00001hjT9RYAU":["engaged","Very interested; visit Saturday with brother","74f82b8ebc57"],"006PW00001hjTMLYA2":["engaged","Noor 131m; site visit next Saturday","5dd2bf28a557"],"006PW00001hjTivYAE":["lost","Paid EOI with competitor; likely going there","2ccf3cfc8245"],"006PW00001hjTxRYAU":["hesit","Postponed; business not stable","fdaa86ade9fc"],"006PW00001hjcZJYAY":["hesit","Travelling; back in 2 weeks","5fedda18350f"],"006PW00001hjdbpYAA":["wait","Waiting specific 123m open-view unit","911ad049000b"],"006PW00001hjeo4YAA":["hesit","Comparing options while travelling","21637a306228"],"006PW00001hjleDYAQ":["engaged","Discussing 119m 3BR; zoom today","58e0af88d958"],"006PW00001hjyUzYAI":["wait","Privado too pricey; waiting B8 release","e734992e67eb"],"006PW00001hk0LVYAY":["unresp","Was to pay DP; now hesitant, not answering","4332d963b688"],"006PW00001hnXsnYAE":["engaged","Reviewing 3BR numbers; will visit branch","dfd7d153ca48"],"006PW00001ho1PSYAY":["engaged","Ready to invest now; funds coming","58365616f587"],"006PW00001ho8yoYAA":["wait","Wants 82m; waiting next release","583b2e603e79"],"006PW00001hoLSzYAM":["hesit","Visited; still studying the market","899991f2c4c5"],"006PW00001hoQNpYAM":["unresp","Wants Noor shop; calls failing","524a89f54b24"],"006PW00001hoQalYAE":["wait","Recent SouthMed buyer; will wait for Alex","301c5161851d"],"006PW00001hoRN8YAM":["unresp","Keeps setting meetings and not showing","63d9185100c9"],"006PW00001hoajlYAA":["wait","Prefers West Cairo","58cffb6792b3"],"006PW00001hpDy2YAE":["wait","Ready to close once Noor 82m releases","1a1291ba8c6b"],"006PW00001hpXIQYA2":["hesit","Travelling; visit next month then decide","ac41dcb0eba5"],"006PW00001hr2jxYAA":["hesit","Back from travel; postponed to new year","0ae3c13910f9"],"006PW00001hrVXRYA2":["unresp","Repeatedly no-shows; not answering","ba92c324c8f1"],"006PW00001hs5XhYAI":["hesit","Back from travel in November","867473867f6b"],"006PW00001htmYzYAI":["none","Only requirement: 3BR Noor","6f7766736e06"],"006PW00001hyKdiYAE":["wait","Waiting Alex project","000330654d57"],"006PW00001hyvYPYAY":["wait","Will bring cheque for 82m at next release","f173e5e195ae"],"006PW00001hzH97YAE":["hesit","Meeting cancelled due to circumstances","cea7ecd9efb7"],"006PW00001hzHlpYAE":["engaged","Ship captain; talk Thursday","3cab43e23408"],"006PW00001hzPgIYAU":["hesit","Choosing RTM villa; checking finances with CFO","87b292fb0e4b"],"006PW00001hzRObYAM":["none","Only requirement: 119m Noor","d9a7bd503d1f"],"006PW00001hzjF4YAI":["hesit","Visited Noor; arranging down payment","ffd5cd4e315e"],"006PW00001hzqgMYAQ":["engaged","Abroad; zoom on 2BR RTM","63e864383536"],"006PW00001hzsNCYAY":["engaged","Meeting done; wants investment options","d1d2d2ef6f78"],"006PW00001i07vhYAA":["engaged","Meeting today, Noor 119m","092d57312a4e"],"006PW00001i0FoPYAU":["engaged","Meeting rescheduled to Saturday","e4cea14806f4"],"006PW00001i0MptYAE":["wait","Slow decision; waiting 116m typical-floor release","83295bdcf886"],"006PW00001i0oSnYAI":["book","Reserved unit 023/038/404","4c0bdefa6052"],"006PW00001i1FO9YAM":["hesit","Comparing 2 vs 3BR; always busy","dd533254c084"],"006PW00001i1FXqYAM":["engaged","Site visit Wednesday, 3BR 123m","ea4a1418b1c3"],"006PW00001i1H09YAE":["engaged","Asked 3BR Noor/SouthMed; details sent","1176510e3e7b"],"006PW00001i1TXaYAM":["unresp","No answer; Sharm Bay 3BR","1c853c6960da"],"006PW00001i2mQ5YAI":["wait","Interested in Alex project","7bc71ce3a5a0"],"006PW00001i2vpyYAA":["wait","No 82m available; waiting next release or Alex","af05f9c53035"],"006PW00001i39JRYAY":["hesit","Selling Mokattam flat first; very interested","02788c29fb0b"],"006PW00001i3eIPYAY":["none","Only requirement: villa/apartment Noor for son","5563a18ee084"],"006PW00001i3vSvYAI":["unresp","Liked 35m office; now not answering","e428d653f1e8"],"006PW00001i57GjYAI":["engaged","Call back; Noor/Madinaty 2-3BR","66df34c1701a"],"006PW00001iAbndYAC":["hesit","Funds in 1.5 months","99c587d5c547"],"006PW00001iAo1iYAC":["engaged","Meeting next week, Privado 2BR","04984747f532"],"006PW00001iAwqTYAS":["none","Only requirement: SouthMed 1BR","0cfccbd7d9a9"],"006PW00001iBNU5YAO":["engaged","Meeting Saturday, 82/98m","76ca776b0e0f"],"006PW00001iBQzpYAG":["hesit","Preparing herself until month end","02c4483038c4"],"006PW00001iBUtlYAG":["wait","Interested in Alex project","b148a96f431e"],"006PW00001iBV3RYAW":["hesit","Abroad; will check needs","db5146db18b8"],"006PW00001iBXV4YAO":["lost","Settled on competitor (Al Kayan); fallback only","a289934b4c4f"],"006PW00001iBjPlYAK":["engaged","Meeting within 3 days, 2BR chalet","053de4f65db0"],"006PW00001iBjxdYAC":["unresp","Selling Shorouk flat first; not answering","af489bde0c04"],"006PW00001iC5WjYAK":["engaged","Came to office; investing for friends","af4e59324702"],"006PW00001iC72HYAS":["engaged","Existing SouthMed buyer; pitching another","fa9017472bd3"],"006PW00001iCE2AYAW":["book","Coming Saturday to close 1BR Golf","5d8c9a6b7649"],"006PW00001iDGd0YAG":["hesit","Was booking 146m; husband's circumstances, postponed","077114e9a7df"],"006PW00001iDzNBYA0":["hesit","Liked Privado; checking resale options","e34ffc7ff2c2"],"006PW00001iE7hJYAS":["wait","Bought 3 Noor units; waiting ground 2BR chalet","562225dd42c4"],"006PW00001iMOTFYA4":["engaged","Zoom done; decides Noor vs SouthMed Saturday","db620dc691a3"],"006PW00001iMTSvYAO":["wait","Wants 2BR in Alex","4272c4f40db3"],"006PW00001iN0ofYAC":["wait","Waiting lake-view Privado 3BR RTM","c9de96ee924b"],"006PW00001iN3EfYAK":["wait","Bringing 82m unit from today's release","e44ce9a62872"],"006PW00001iO0Y5YAK":["engaged","Settled on small investment unit","d5b2d81e65c1"],"006PW00001iPWu7YAG":["lost","Not interested; in Dahab","ce8e8e2fc10c"],"006PW00001iQEiDYAW":["lost","Met in Sharm; not a buyer","7b2d84d1e073"],"006PW00001iaPZ7YAM":["engaged","Zoom with husband; liked 145m, price worry","5e1057e5a722"],"006PW00001if6vdYAA":["hesit","Financial problem; proceed in 2 months","4a5f0332cc1d"],"006PW00001iidS1YAI":["engaged","Choosing unit; budget not final","224de5adb3e4"],"006PW00001iidbhYAA":["engaged","Meeting today, Noor 2BR","e129394eed6f"],"006PW00001iijcIYAQ":["hesit","Postponed to December","223e180f93f4"],"006PW00001iilnmYAA":["hesit","Son travelling; will execute mid-November","b6efe440e40b"],"006PW00001iioKFYAY":["hesit","Robbed on arrival; sorting papers, travelling","ef4e7da8a0eb"],"006PW00001iiqFZYAY":["hesit","Chose unit; not ready, room-size concern","91e5dc8902f7"],"006PW00001iiqXJYAY":["wait","Asking about Alex","12ec09810bae"],"006PW00001iiuPdYAI":["wait","Waiting Alex project","89ee76f4828a"],"006PW00001iiuSrYAI":["hesit","Financial problem; money in November","bf6e9ae45a4c"],"006PW00001iiwuTYAQ":["hesit","Missed meeting; circumstances","7aeabb53221e"],"006PW00001ij0QDYAY":["hesit","Not ready; will come on vacation","1e8cbcf72517"],"006PW00001ij0wTYAQ":["engaged","Following on WhatsApp; explained location","00ccce26e5f0"],"006PW00001ij11JYAQ":["engaged","3BR investment; coming to Cairo","e896f69b12fe"],"006PW00001ijFM1YAM":["engaged","Meeting Thursday for son's quattro","836fc2cd29a1"],"006PW00001ijGWbYAM":["unresp","Zoom planned; no answer","f820fa376983"],"006PW00001ijWrWYAU":["engaged","Visited Privado; arranging Noor quattro visit","dcb82455f84f"],"006PW00001ijZ7sYAE":["wait","Waiting Alex project","5cde64e9cb1b"],"006PW00001ijf5BYAQ":["hesit","Abroad in Riyadh; will follow up","4b3e6e7a49c1"],"006PW00001ijtL3YAI":["wait","Waiting Alexandria project","59d9f67b3bc9"],"006PW00001ik1fBYAQ":["engaged","Meeting tomorrow with husband; prices high","da2a3f73b845"],"006PW00001ik41zYAA":["engaged","Coming tomorrow with husband; annual-payment plan","49c1ccef2db1"],"006PW00001ikhmVYAQ":["wait","Waiting Noor 1BR release","06889c3a66dd"],"006PW00001il0qmYAA":["none","No notes","da39a3ee5e6b"],"006PW00001iofCcYAI":["hesit","Undecided Noor vs SouthMed","22286af57e76"],"006PW00001ioilZYAQ":["wait","Thought Noor was in Alex; may wait","5e773c3ba4fe"],"006PW00001iojksYAA":["wait","Working with another rep; next Noor release","c4e0c4e1e7d3"],"006PW00001iokiYYAQ":["hesit","Comparing Privado vs B14; needs time","f4292972ad4a"],"006PW00001iopgbYAA":["hesit","Liked 131m; unsure when she visits Egypt","244a0077595c"],"006PW00001ioqZSYAY":["book","Coming tomorrow to book a unit","01a6a7d079a9"],"006PW00001iorLpYAI":["wait","Shop too pricey; waiting cheaper shops","494f39ce1ab4"],"006PW00001iorNRYAY":["engaged","Doctor, hard to reach; following","73f31b5e429c"],"006PW00001ioshhYAA":["unresp","Waiting for reply","32138acb9db2"],"006PW00001ioszRYAQ":["hesit","Prices fine; postponed 2 months to visit","e973c4fc43cb"],"006PW00001iotXJYAY":["wait","Wants Noor 67m","f5a74e6a4d1a"],"006PW00001iovPUYAY":["hesit","Will buy 98m end of year","0e95c7f4ee86"],"006PW00001ioxhNYAQ":["hesit","Settled on 200m; asked for time","9b66d59d47f1"],"006PW00001ip36sYAA":["hesit","Very interested; ill, meeting next week","f8faf9862010"],"006PW00001ip38TYAQ":["engaged","Wants to buy Noor quattro; hard to schedule","85e905c5f60d"],"006PW00001ip3gLYAQ":["wait","Waiting for 67m or 82m","a4c1acec219e"],"006PW00001ip8xmYAA":["hesit","Will buy in a month","7b86231a7dbb"],"006PW00001ipBFhYAM":["hesit","Very interested; not before December","f8f43e08b3b8"],"006PW00001ipG0rYAE":["hesit","Settled on lake-view Privado; too expensive","9eb91330194c"],"006PW00001ipLoYYAU":["hesit","Travelling; back in November","1193d6715573"],"006PW00001ipcxRYAQ":["engaged","Site visit Tuesday being arranged","5ae8ff5d3d0b"],"006PW00001ipgONYAY":["engaged","Zoom Thursday 8pm, B14","4b0219462bc6"],"006PW00001iqHniYAE":["none","Only requirement: 3BR 119m","9b519d312b50"],"006PW00001iqjIYYAY":["engaged","Wants Four Seasons villa; meeting tomorrow","9afb4e0fc53f"],"006PW00001iqpfjYAA":["engaged","Zoom done; reply within a week","ee425c469e4b"],"006PW00001iscwAYAQ":["hesit","Price too high; checking Noor","2440d13d4eb7"],"006PW00001ishe3YAA":["lost","Lost interest","5f4ec442cf7e"],"006PW00001it9GwYAI":["engaged","Details sent, 3BR 131m","81ab9cc5d045"],"006PW00001itCQ5YAM":["wait","Wants 82m to book; waiting new release","c27e0edf2c16"],"006PW00001itCWXYA2":["hesit","Chose Golf 2BR; postponed, sister needed money","2285cee09784"],"006PW00001itP2LYAU":["hesit","Undecided","bd15e148830b"],"006PW00001iu03VYAQ":["engaged","Wants cemetery plot; site visit set","8ef58c73ed9a"],"006PW00001iuAVtYAM":["book","Liked Privado; deciding today on booking","bf798cad5e38"],"006PW00001iuRrjYAE":["none","Only requirement: Noor 82/119m","048c97517ffa"],"006PW00001iueSLYAY":["hesit","Very interested; will proceed 1 Dec","71a6089d9ba7"],"006PW00001iujVGYAY":["unresp","Planned meeting missed; no answer","5895efd7740e"],"006PW00001ivGNxYAM":["book","Awaiting transfer for 82m","cee1891fa89e"],"006PW00001ivJgnYAE":["wait","Repeat buyer; likely Alex project","9aa13f873258"],"006PW00001ivcMrYAI":["hesit","Awaiting Kuwait financing approval","d991ea567baf"],"006PW00001ivdkMYAQ":["engaged","Meeting tomorrow, Noor quattro","863f2aa2a1a8"],"006PW00001ivrc3YAA":["engaged","Visited; liked 122m ground unit","ac325d3c58a2"],"006PW00001iw5zxYAA":["hesit","Decision after Ramadan","6008c122a6f4"],"006PW00001iwIPKYA2":["wait","Studio in Noor; waiting next release","994a483e5f12"],"006PW00001iwoFWYAY":["wait","Waiting Alexandria project","84762c56efd0"],"006PW00001izopFYAQ":["none","Only requirement: Spine office, 16-25M","7ac07c7bebd9"],"006PW00001j0eJBYAY":["hesit","Almost booked; finishing concern, visiting in 12 days","dac080288ceb"],"006PW00001j4MCQYA2":["hesit","Must sell her flat first","2ce2d2e1d683"],"006PW00001j4as5YAA":["engaged","Met at Dokki; residency issue","b98de75eeb26"],"006PW00001j4jgrYAA":["engaged","Wants Noor investment; will review offers","d05284bbb27a"],"006PW00001j4nkTYAQ":["engaged","Meeting today, Privado 2BR RTM","f8526e7f02f9"],"006PW00001j4ymNYAQ":["hesit","Selling her flat first; just browsing","5ceb3cfa39a8"],"006PW00001j5BWgYAM":["hesit","Travelling; no step until November","e70bcc02bb6b"],"006PW00001j5nX7YAI":["unresp","Consulting family; not answering","00d6fd9c9738"],"006PW00001j5wgrYAA":["engaged","Coming Thursday with son","8ff052ec9695"],"006PW00001j5xUrYAI":["unresp","Waiting lake-front RTM; went silent after zoom","f0ee9ad8931a"],"006PW00001j68q5YAA":["engaged","Arranging Privado RTM meeting","95c2fe72e504"],"006PW00001j69uDYAQ":["engaged","Will schedule visit this week","429fcb52d36d"],"006PW00001j6KfxYAE":["none","No notes","da39a3ee5e6b"],"006PW00001j6MBVYA2":["lost","Not interested","d733598cad13"],"006PW00001j6lD7YAI":["hesit","Gave cheque at Kayan City; visiting Noor Saturday","d8b656055c63"],"006PW00001j6pjlYAA":["engaged","Coming early Nov to view 98m","ffc1b0ea9e54"],"006PW00001j6uecYAA":["engaged","Site visit done; negotiating Privado 2BR","f710f1c4482c"],"006PW00001j707JYAQ":["engaged","Setting meeting this week","8c6fa3c846c4"],"006PW00001j7YH3YAM":["unresp","Studying Lagoon/Golf plans; now not answering","9d34941e840e"],"006PW00001j7evxYAA":["engaged","Interested in Sharm, maybe two units","e51a189f9734"],"006PW00001j7h7SYAQ":["hesit","Postponed two months","3115528d3717"],"006PW00001j806rYAA":["hesit","Factory problems (dollar); not ready","d9ff36ff35a1"],"006PW00001j80ExYAI":["engaged","Interested; following on WhatsApp","7d634d0e1889"],"006PW00001j8GgHYAU":["hesit","Still praying on it; undecided","9f47979ef9c7"],"006PW00001j8J7tYAE":["hesit","Travelling 2 weeks","594882216166"],"006PW00001j8JuHYAU":["hesit","Choosing 98 vs 119m; checking bank limit","bc73388b4d38"],"006PW00001j8NRdYAM":["unresp","Considering Kayan; not answering","5881b4664f63"],"006PW00001j97z9YAA":["wait","Waiting 1BR release; thinking about EOI","a505af883359"],"006PW00001j9pfBYAQ":["hesit","Settled Privado RTM; consulting children abroad","bbd1a5417552"],"006PW00001jBMqnYAG":["hesit","Comparing with Kayan City; still thinking","fd7da161bb67"],"006PW00001jBOW1YAO":["hesit","Abroad in KSA; back in 2 weeks","a449e7c46e2e"],"006PW00001jBg1VYAS":["wait","Waiting 82m in Noor","4e48574af7ef"],"006PW00001jFaDfYAK":["wait","Waiting 82/67m in next Noor release","f3b792d508e7"],"006PW00001jFk4wYAC":["wait","Interested Noor 1BR 67m","cec5b3666c75"],"006PW00001jH1tJYAS":["hesit","Back early next month","9151ba00aef2"],"006PW00001jH3v8YAC":["none","No notes","da39a3ee5e6b"],"006PW00001jHnWXYA0":["hesit","Going to Umrah; call 11-12","eb3edbfafd10"],"006PW00001jHtvKYAS":["engaged","Interested 1BR garden Noor","4a88198cd6db"],"006PW00001jI1UfYAK":["engaged","Coming tomorrow 4pm","5add540a6144"],"006PW00001jI4ynYAC":["wait","Alex project villas 3-5BR","96c1d961f9fd"],"006PW00001jI9f8YAC":["book","May close tomorrow or after","5938ca7b72f3"],"006PW00001jIDysYAG":["hesit","Travelling to Germany 2 weeks","410ff2d7fd3a"],"006PW00001jIYyUYAW":["wait","Waiting resale/cancelled ground 2BR Lagoon","2f9aaaeaea64"],"006PW00001jIzKLYA0":["wait","Interested in Alex project","0451c216ef50"],"006PW00001jJ1POYA0":["wait","Wants Noor 2BR; waiting release","4c4c1ce27d07"],"006PW00001jJF7OYAW":["engaged","Comparing Noor and Privado","1e87ff17ea38"],"006PW00001jJR0TYAW":["unresp","Rescheduled twice; no answer","5d56356065f1"],"006PW00001jJZfaYAG":["hesit","Out of Cairo; can't zoom","80ba6f8d45cb"],"006PW00001jJi69YAC":["engaged","Officer in Sinai; asking new Madinaty project","b313b81e02f7"],"006PW00001jJkeDYAS":["hesit","Retired; buy sometime in future","2acef3f42075"],"006PW00001jK09VYAS":["engaged","Coming to view 146m lake unit tomorrow","f6a16fea7095"],"006PW00001jLa2rYAC":["wait","Interested 3BR Alex project","fe5367193d81"],"006PW00001jRx82YAC":["engaged","Meeting Tuesday, Noor 2BR","aa7d62696309"],"006PW00001jS99DYAS":["wait","Thinking on offer; waiting Alex project","b48585364f74"],"006PW00001jSDJMYA4":["wait","Eagerly waiting Alex launch","5e9bd6ee2406"],"006PW00001jSPylYAG":["engaged","Setting Privado 2BR meeting","c051fc9feabc"],"006PW00001jSWVZYA4":["none","Only requirement: 3BR ground garden","bc5f256c4be4"],"006PW00001jSZ21YAG":["wait","Group buyers; want lake-view RTM Privado","a55a215ef520"],"006PW00001jSZtFYAW":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00001jSe1hYAC":["lost","Bought via another broker (conflict check)","8a0e9101eedd"],"006PW00001jSgGPYA0":["wait","Eagerly waiting Alex release","06ffd2906cc6"],"006PW00001jSgRhYAK":["hesit","Thinks rooms too small; reply in 2 weeks","16fcb3419a79"],"006PW00001jSqJ0YAK":["wait","Alexandria only","cc9c5792aaf2"],"006PW00001jSvfHYAS":["wait","Chosen units sold; waiting new release","3547e8a9df26"],"006PW00001jTUzbYAG":["hesit","Unsure about finances; leans Abu Qir","c31cfd002f77"],"006PW00001jTuZ4YAK":["hesit","Comparing resale offers","b0f24c473bce"],"006PW00001jUVWzYAO":["hesit","Travelling; call 1 Nov","5f5ebee0fc8d"],"006PW00001jVeWcYAK":["wait","Wants RTM overlooking Spine","00d118b7d88e"],"006PW00001jVzw1YAC":["book","Viewing alternative unit to contract","b7af356b3b0a"],"006PW00001jXB0fYAG":["book","Back in touch; will reply and transfer","20b54f12c0ba"],"006PW00001jXrEPYA0":["wait","Waiting Alex project","322260d8ebfb"],"006PW00001ja56nYAA":["wait","Asking about Alex projects","3c5ae42a167b"],"006PW00001jesSKYAY":["unresp","Picked 2 units on zoom; silent since","04e19245c7df"],"006PW00001jf1VZYAY":["hesit","Funds in 10 days; then 1BR La Cantine","3c3c5b17ca67"],"006PW00001jgYNpYAM":["engaged","Choosing best location; zoom before visit","53aebc886592"],"006PW00001jjxnwYAA":["engaged","Wants all project details","c2d56472c7e9"],"006PW00001jlB2PYAU":["engaged","Will call back in 10 minutes","d50f593c8638"],"006PW00001jlDDtYAM":["none","Follow up on WhatsApp only","633076c648a8"],"006PW00001joN3qYAE":["unresp","Broker; not answering","8cc1a2ef0975"],"006PW00001johNYYAY":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00001jojacYAA":["wait","Interested in North Madinaty project","6ef1ae3153f0"],"006PW00001joqIkYAI":["wait","Waiting Alex project","f13622dd6119"],"006PW00001jorZmYAI":["wait","Waiting Alex details","860c6281fd26"],"006PW00001jot8XYAQ":["engaged","Meeting done, 1BR Lagoon","1125bec7f156"],"006PW00001jovOtYAI":["engaged","Coming to Cairo early next month","d5b1bf2b74cc"],"006PW00001jovv7YAA":["wait","Interested in Alex project","1abd6484c514"],"006PW00001jowkjYAA":["unresp","Number unreachable","93126dd12a89"],"006PW00001joyRZYAY":["hesit","Husband travelling; decide after 9th","bb6da718ea29"],"006PW00001joz5uYAA":["engaged","Coming Saturday; 2BR Lagoon primary","6c6869255f7a"],"006PW00001jozSTYAY":["lost","Lost interest","dc66a1db6246"],"006PW00001jp00LYAQ":["wait","Waiting Alex project","d06c67677d68"],"006PW00001jp0MvYAI":["engaged","Reviewing SouthMed 1BR payment","d9494e0da647"],"006PW00001jp2YPYAY":["hesit","Liked; payment concerns, back mid-Nov","1a2bf76a0e14"],"006PW00001jp2uzYAA":["wait","Alex project; follow up at details","227be189e6e2"],"006PW00001jp4FHYAY":["hesit","Might travel; may drop idea","8fce8cf7102d"],"006PW00001jp4ITYAY":["none","Only requirement: 3BR Madinaty","6b9493b5f884"],"006PW00001jp56TYAQ":["engaged","Husband in UAE engaged; following","0a7b40f81a14"],"006PW00001jp8E1YAI":["hesit","Arranging money; working with another rep","f14598cd8c62"],"006PW00001jpFaTYAU":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00001jpIDPYA2":["wait","Wants typical 3BR >140m; waiting next release","c835936b4933"],"006PW00001jpP6nYAE":["wait","Interested in Alex","d8df4726d1d2"],"006PW00001jpPV0YAM":["engaged","Zoom set, Sharm Bay 1BR","be7a204c7179"],"006PW00001jpPWbYAM":["hesit","Committed elsewhere; maybe April","4f940ed90e65"],"006PW00001jpagYYAQ":["wait","Wants 1BR 67m when released","b71d548c9bb7"],"006PW00001jpe4FYAQ":["hesit","Circumstances; meeting early November","049f94523985"],"006PW00001jpp2rYAA":["hesit","Travelling; talk next week","697803ba81ed"],"006PW00001jppR3YAI":["wait","Repeat buyer; waiting Alex details","70d976af05b3"],"006PW00001jpqDRYAY":["engaged","Calculating Noor townhouse payment","17a101416959"],"006PW00001jpuiVYAQ":["hesit","Busy this week; next week","323eafc1f827"],"006PW00001jq40JYAQ":["engaged","Abroad; Noor site visit to be set","ceeedc98e793"],"006PW00001jq7B3YAI":["engaged","Site visit done, 146/147m","e8ecca6b7bd0"],"006PW00001jqGW7YAM":["engaged","Zoom done; offer being sent","58c3b45af26d"],"006PW00001jqVrhYAE":["hesit","Slow decision-maker; long way off","324f3f489e44"],"006PW00001juA8kYAE":["wait","Asking Alexandria details","c8a4aeb99a3e"],"006PW00001juEPHYA2":["wait","Waiting Alex launch","3821cbfc92b7"],"006PW00001juf9MYAQ":["engaged","Zoom done; near-delivery Privado/Madinaty","7f6b3f314201"],"006PW00001juhhNYAQ":["wait","Waiting 98m unit under 8M","cceee112ec33"],"006PW00001jupILYAY":["wait","Wants Abu Qir; not answering","8f434a34e64d"],"006PW00001juqRJYAY":["wait","Interested Alex apartment","3ff34586dab8"],"006PW00001jur0nYAA":["unresp","Liked 123m; postponed, now no answer","0de3e5962d77"],"006PW00001jur8sYAA":["wait","Interested in Alex projects","bb504f1abc71"],"006PW00001jusg1YAA":["engaged","Requested 2-3BR details","7aaba1e70f8b"],"006PW00001juvSbYAI":["wait","Wants Alex project, not Sahel","94c29d872fd9"],"006PW00001juwenYAA":["engaged","Wants RTM, no annual payments; 3BR Noor","24c53d31a4ff"],"006PW00001juznyYAA":["engaged","Will try to come this week","e95e4ec5ca29"],"006PW00001jv13NYAQ":["engaged","Noor owner; reviewing 119m payment","dd5b5f9f675b"],"006PW00001jv1EfYAI":["wait","Waiting Alex launch","eaf7bfad10b3"],"006PW00001jv58cYAA":["wait","Wants 2BR in Alex","52b22b452306"],"006PW00001jv5jhYAA":["hesit","Back on 30 Nov","b6088a2ec8de"],"006PW00001jv8mQYAQ":["unresp","Prefers Giza; not answering","789be469a5c6"],"006PW00001jv9vNYAQ":["wait","Interested in Alex only","60c6d277a8bd"],"006PW00001jvAzVYAU":["wait","Wants 2BR in Alex","2afe7e9f0ac0"],"006PW00001jvC0PYAU":["wait","Waiting release","f82561c73c8a"],"006PW00001jvEa5YAE":["engaged","Wants RTM 2BR; Privado sent","54296336f70d"],"006PW00001jvG7GYAU":["wait","Low budget; waiting Alex","4dfdd43c1f0a"],"006PW00001jvGbtYAE":["wait","Asking about Alex","c8eae14c8da8"],"006PW00001jvNaAYAU":["hesit","Selling his other flat first","4a75f63ee40b"],"006PW00001jvRPHYA2":["hesit","Dollar funds delay; repeated no answer","3d67dad605fb"],"006PW00001jvSpxYAE":["wait","Wants Alex","2a8ada4b1caf"],"006PW00001jvZzVYAU":["wait","Wants 180m Privado/B8; searching unit","886d9fffa7f6"],"006PW00001jvgXyYAI":["wait","Waiting Alex project","698d88b0997c"],"006PW00001jvjKXYAY":["book","Will transfer from abroad; waiting passport","2d4cc2451154"],"006PW00001jvsm1YAA":["engaged","Meeting next week, 82m","1d8ed72fbdaf"],"006PW00001jvwmSYAQ":["wait","Call when Alex launches","15472e5c35b6"],"006PW00001jwC3CYAU":["wait","Asking about Alex project","527efdfcda16"],"006PW00001jwOM9YAM":["wait","Waiting Alex project","d27704af5480"],"006PW00001jyKmEYAU":["engaged","Visited; comparing 131 vs 146m","dd5fb862f474"],"006PW00001jz0tVYAQ":["wait","Call when info available","463022870095"],"006PW00001jz9FGYAY":["hesit","Selling a flat; a month away","04e60c755c14"],"006PW00001jzCUrYAM":["engaged","Asking about new Privado buildings","1af3ee0f85fb"],"006PW00001jzIK9YAM":["unresp","Not answering; last attempt","8e6515daa314"],"006PW00001jzKXFYA2":["wait","Interested in Alex project","51022226dd13"],"006PW00001jzKx3YAE":["engaged","Studying Privado 3BR with husband","4d18a176f8db"],"006PW00001jzNBlYAM":["engaged","Liked resale chalet; meeting at his office","cab5034414cd"],"006PW00001jzq8vYAA":["wait","Will pay EOI before release","8e59deee6eda"],"006PW00001k00jNYAQ":["engaged","Awaiting reply on last 145m ground unit","42d96f862af0"],"006PW00001k0HnRYAU":["wait","Waiting Alex launch","7caf842479ab"],"006PW00001k0JCZYA2":["hesit","Not available to execute now","7ffc501b0be0"],"006PW00001k0JKbYAM":["wait","Waiting Alex launch","ad3d0f90e498"],"006PW00001k0L6HYAU":["book","Upgrading from 119m to 131m tomorrow","cfb9cc771892"],"006PW00001k0M3xYAE":["wait","Waiting Alex project","7b4ff5075766"],"006PW00001k0b7qYAA":["none","Only requirement: 3BR B14/Noor/Privado","190693993170"],"006PW00001k0bMMYAY":["book","Will transfer booking renewal","aeddd4ad1c90"],"006PW00001k0crtYAA":["engaged","Discussing Sharm Bay 1BR on WhatsApp","1eab123be086"],"006PW00001k0fg5YAA":["wait","Interested in Alex only","ed334de8558e"],"006PW00001k0lYbYAI":["engaged","Discussing 98m installments","6fbc9fd28adc"],"006PW00001k17MDYAY":["hesit","Busy abroad; may hold a month","6a4c84a7a741"],"006PW00001k1G6FYAU":["none","Only requirement: Noor","a0876dea9117"],"006PW00001k1Xi6YAE":["wait","Waiting Noor 67m release","3ed7153632b0"],"006PW00001k1cOSYAY":["hesit","Meeting start of month when funds arrive","915e75ee6a5e"],"006PW00001k1j8BYAQ":["wait","Interested in Alex project","db53ecbca755"],"006PW00001k1oPaYAI":["engaged","Meeting done, ground garden Golf","2ac1fab86232"],"006PW00001k1oSnYAI":["hesit","Not ready until new year","95d7cbec196e"],"006PW00001k1sjJYAQ":["book","Came to renew booking for 82m","34f9a6c8baa1"],"006PW00001k1t5uYAA":["wait","Interested in Alex project","0e3a652da1a5"],"006PW00001k1uJhYAI":["engaged","Will set meeting this week","a3c1bfff3b55"],"006PW00001k1wVBYAY":["engaged","Will visit","01c5768a2871"],"006PW00001k1zkpYAA":["wait","Waiting Alex launch","7caf842479ab"],"006PW00001k2A1uYAE":["book","Visited; left EOI for Wednesday release","f0b7f3542ed0"],"006PW00001k2BdtYAE":["unresp","Comparing competitor; cancelled, not answering","f90a884c3c32"],"006PW00001k2XJRYA2":["engaged","Meeting at 6pm, Noor 82m","d1c846254958"],"006PW00001k2aArYAI":["engaged","Meeting Saturday, Sharm 1-2BR","265ce8975711"],"006PW00001k2gJYYAY":["engaged","Changed to Privado 1BR","7f1c6b2b9a78"],"006PW00001k36VjYAI":["hesit","Bereavement; meeting rescheduled","e25b224733ef"],"006PW00001k3FtzYAE":["engaged","Coming Monday, 82m Noor","d591f441360f"],"006PW00001k3UMjYAM":["engaged","Coming Wednesday/Thursday after travel","75a328cdf741"],"006PW00001k3m58YAA":["engaged","Will buy one 123m unit","f3accbc2955c"],"006PW00001k4EB4YAM":["wait","Asking about Alex project","26b38887a40f"],"006PW00001k4IjKYAU":["hesit","Comparing prices; reply in 2 days","d300c5306b73"],"006PW00001k4NO1YAM":["unresp","Prefers Privado RTM; site visit unsettled, no answer","46682184e2a3"],"006PW00001k4PEXYA2":["engaged","Reviewing Lagoon/Marina payment","dbebbc0c5b53"],"006PW00001k4PzJYAU":["engaged","Interested in quattro; call Monday","1d89b99b1e57"],"006PW00001k4QovYAE":["unresp","Agreed on 3BR 5-yr plan; phone off","372838d26cb1"],"006PW00001k4YXxYAM":["book","Paid booking deposit; waiting new Noor release","27c4da0600aa"],"006PW00001k53wjYAA":["engaged","Wants two units; coming next week","835670bafe5c"],"006PW00001k56Y2YAI":["none","Only requirement: 2BR Noor","41d73a022a6d"],"006PW00001k6GiDYAU":["lost","Cash buyer of resale only; bought 3 units elsewhere","5097c4a5cb5b"],"006PW00001k77EfYAI":["none","Only requirement: SouthMed villa/3BR","a7b5b6266622"],"006PW00001k79A3YAI":["wait","Waiting Alex project","033834b8c37e"],"006PW00001k7mRWYAY":["engaged","Studying 67m plan with brother","46374632c3fb"],"006PW00001k8VDKYA2":["book","EOI for Alex apartment release","85a1dc4caffa"],"006PW00001kDFV4YAO":["hesit","Liquidity issue; call 15 Nov","04150cdd1c45"],"006PW00001kDS1wYAG":["unresp","Bereavement; meeting rescheduled, no answer","b9219ef07349"],"006PW00001kDc2tYAC":["wait","Interested in Alexandria","29d2b9bf06e5"],"006PW00001kDdvNYAS":["engaged","Zoom done; comparing 131 vs 146m","6896f04b321c"],"006PW00001kDmKgYAK":["hesit","Meeting postponed; husband busy","aae2301ab634"],"006PW00001kDscIYAS":["book","Will pay 50K for 119m at new launch","ec7610a1be10"],"006PW00001kDzIoYAK":["wait","Waiting Alex project","22de644dbd7b"],"006PW00001kE0LJYA0":["engaged","Buying 98m with two friends; consulting them","f2cf3751ea4d"],"006PW00001kE0y4YAC":["wait","Interested in Alexandria","1ccb283eff0e"],"006PW00001kE2F3YAK":["engaged","Visiting Noor and SouthMed sites with friend","a7dcce35ccba"],"006PW00001kE3avYAC":["unresp","Planned office visit; no answer since","6365fa6e96ac"],"006PW00001kE3fmYAC":["wait","Wants Alex; phone closed","88068bac453a"],"006PW00001kE71pYAC":["engaged","Considering payment; reply Thursday","e675b77e9c05"],"006PW00001kEEBNYA4":["wait","Wants Alex details when out","40d1d6069750"],"006PW00001kEF0zYAG":["engaged","Comparing Noor townhouse vs resale","caa350f27e1a"],"006PW00001kEF4DYAW":["wait","Wants Alex project","99acfca2bdc5"],"006PW00001kEFFVYA4":["book","Coming tomorrow with cheques for SouthMed","26e5c50b9a8b"],"006PW00001kEFH7YAO":["hesit","Interested, but in January","e9e75614c1d2"],"006PW00001kEGjRYAW":["engaged","Office meeting next week","7e18a27f0ce3"],"006PW00001kEJkXYAW":["engaged","Meeting","bd7580126b94"],"006PW00001kEPTNYA4":["hesit","Selling Noor 82m first, then villa","8bb88a6fc4b1"],"006PW00001kEVddYAG":["engaged","Referral; Noor site visit Sunday","51675fc380bc"],"006PW00001kEX7ZYAW":["wait","Waiting Alex project","9fdc52cd672e"],"006PW00001kEhYLYA0":["unresp","No answer","9d1ad0b5b25a"],"006PW00001kEn2gYAC":["wait","Wants Alex project","a99240abc4e7"],"006PW00001kF0WQYA0":["book","Will join allocation tomorrow; 3 units 123m","7db1c0d1fb47"],"006PW00001kF5qnYAC":["engaged","Discussing Noor on WhatsApp","2bf1456c38bd"],"006PW00001kF7B3YAK":["engaged","Asking about B14 and Noor","82edc7480202"],"006PW00001kF7CfYAK":["wait","Wants Alexandria or October","46093ed166f6"],"006PW00001kF7cUYAS":["engaged","Asking for office visit","af3fe86a1640"],"006PW00001kFRL3YAO":["hesit","Waiting husband's return end of month","ab49f28a6480"],"006PW00001kFStqYAG":["engaged","Meeting Tuesday, Spine 1BR","2965e832f808"],"006PW00001kFVq7YAG":["none","No notes","da39a3ee5e6b"],"006PW00001kFcV0YAK":["hesit","Travelling; visit on return","48e1f6fe0b87"],"006PW00001kFn4rYAC":["hesit","Still deciding with her mother","c393b02fc37c"],"006PW00001kFzBRYA0":["engaged","Family visit to Noor after Saturday","5a07988a9907"],"006PW00001kFzEhYAK":["hesit","Back in 4 months; refused zoom","61db4cdd9557"],"006PW00001kG9M6YAK":["hesit","Selling Noor flat to buy Privado","7ea782d405ff"],"006PW00001kGLThYAO":["wait","Waiting Alex release","1be7f1b33749"],"006PW00001kGQOXYA4":["wait","Waiting good-location Noor 2BR","5013d625cd93"],"006PW00001kGW5mYAG":["engaged","Liked prices; will visit to choose location","b3b3c2aa8e55"],"006PW00001kGbn1YAC":["engaged","Office visit Sunday/Monday","393a0257694c"],"006PW00001kGdK9YAK":["wait","Waiting Alex launch","a870c1379247"],"006PW00001kGfLyYAK":["unresp","Not answering; in Alex on business","e3ac786aa800"],"006PW00001kGg6jYAC":["none","Only requirement: 119/146m Noor","53f50aaf61c2"],"006PW00001kGiIDYA0":["wait","Waiting affordable 1BR","a5735a9e676f"],"006PW00001kHGRzYAO":["book","Ready to buy within a week","897825e8a90f"],"006PW00001kHLD7YAO":["engaged","Unit picked; calculating affordability","08443bdff43c"],"006PW00001kHMynYAG":["hesit","Not ready at these prices","72ee68abe62f"],"006PW00001kHVXTYA4":["engaged","Privado too pricey; reviewing Noor 2BR","8f3a8113f39b"],"006PW00001kHYDZYA4":["wait","Asking North Rehab 3BR","e9aa2192d74d"],"006PW00001kHlCPYA0":["engaged","Interested in Noor 82m","6639de013c16"],"006PW00001kHuikYAC":["engaged","Zoom at 6, near-delivery units","e15cadf5c2cc"],"006PW00001kIZSXYA4":["wait","Waiting Alex villas","7da17945a642"],"006PW00001kJD9pYAG":["engaged","Thinking on 123m; site visit if OK","d7d16fab4e7e"],"006PW00001kJqESYA0":["none","Only requirement: commercial","1431e668328b"],"006PW00001kL6kBYAS":["wait","Interested in Alexandria","68e5658e70be"],"006PW00001kL6nNYAS":["wait","Alex project","24c7091e3467"],"006PW00001kL7ZoYAK":["engaged","May take two 82m units","a27714bda06c"],"006PW00001kL8AsYAK":["none","Only requirement: 2BR Noor","c7c74668dcf6"],"006PW00001kLwNmYAK":["wait","Son waiting Alex release","471c6823bec1"],"006PW00001kMow2YAC":["none","Abroad; asked 3BR Noor prices","b3afcb56a6e5"],"006PW00001kObLEYA0":["wait","Asking about Alex","cb0e3db8c5e9"],"006PW00001kPi9KYAS":["hesit","Selling assets; ready end of December","0dd4950c9de6"],"006PW00001kSzyRYAS":["book","May close today after husband's ok","60cc513e7c5e"],"006PW00001kT3SZYA0":["unresp","Sent photos; waiting reply","f24a1ec50daa"],"006PW00001kTpAbYAK":["engaged","Arranging Noor visit","4708e0822a7f"],"006PW00001kU4XpYAK":["wait","Wants 2BR in Alex","4d41c76a32fc"],"006PW00001kUDODYA4":["lost","Cancelled Noor; going with Abu Qir","e6a4facce22f"],"006PW00001kUHm2YAG":["engaged","EOI with competitor; second zoom tonight","e455b1fa1f6a"],"006PW00001kULtdYAG":["engaged","Reviewing 98m 2BR; new buildings sent","3c2b11e2a608"],"006PW00001kUMUjYAO":["wait","Asking about Alex project","34954f4118af"],"006PW00001kUPXRYA4":["wait","Waiting Alex release","a49f35cb9a42"],"006PW00001kUSofYAG":["unresp","Meeting Thursday; not answering","32fa389b6002"],"006PW00001kUUQgYAO":["wait","Waiting Alex launch","1124b570266a"],"006PW00001kUYUHYA4":["wait","Waiting Alex launch","1124b570266a"],"006PW00001kUZVBYA4":["wait","Asking about Alex project","58a2a30c4543"],"006PW00001kUbFFYA0":["hesit","Waiting funds end of November","7d22cd11f2b6"],"006PW00001kUds9YAC":["hesit","Not in this release; work issues","78b5516d525a"],"006PW00001kUfqjYAC":["engaged","Coming today to see availability","2c24d8e74b62"],"006PW00001kUgq1YAC":["book","Coming tonight to buy SouthMed","39ee3c09885e"],"006PW00001kUqUTYA0":["none","Only requirement: Madinaty twin house","fa2e63976573"],"006PW00001kUrn8YAC":["hesit","Travelling; visit the week after","270a7bf77bae"],"006PW00001kUxSgYAK":["wait","Repeat buyer; waiting Alex launch","0dd9848ec8c3"],"006PW00001kV4AnYAK":["hesit","Can't come; brother will follow up","cc3b55f1b818"],"006PW00001kV5lCYAS":["none","Only requirement: Noor 1BR garden","2001761cea91"],"006PW00001kVBNZYA4":["engaged","May come Monday for meeting","8cf30d17ecbe"],"006PW00001kVLJhYAO":["wait","Waiting Alex 3BR details","8dacb4353b45"],"006PW00001kVTcDYAW":["wait","Still interested; waiting Alex details","92d0e2edf799"],"006PW00001kVbujYAC":["none","Only requirement: 2BR Noor","1abe1295260c"],"006PW00001kVfDZYA0":["wait","Will wait for Alex project","1a400953c839"],"006PW00001kVs9BYAS":["engaged","Interested 2BR investment Privado","f6b4e7619e3d"],"006PW00001kVti0YAC":["engaged","Arranging office meeting with mother and brother","b3f4d4e229dd"],"006PW00001kW6wxYAC":["engaged","Very interested; visiting end of November","32cc7a4bbfcf"],"006PW00001kW8vWYAS":["wait","Waiting Alex project","f13622dd6119"],"006PW00001kW9oLYAS":["wait","Wants Alex 3BR when info out","763890894728"],"006PW00001kWOk9YAG":["wait","Waiting Alex project","c3be11916e0d"],"006PW00001kWgcFYAS":["wait","Met; waiting new Noor 82m release","f265962d20fb"],"006PW00001kWvwEYAS":["book","Will buy 96m ground garden unit","cf31283c9c81"],"006PW00001kX3VbYAK":["engaged","Site visit Saturday to proceed","7c8397bd3085"],"006PW00001kXHezYAG":["wait","Wants 67m; EOI when buildings open","d9d3900dc58f"],"006PW00001kXJNRYA4":["engaged","Asking 82m/67m payment; release Wednesday","a53083465af3"],"006PW00001kXljVYAS":["wait","1BR in Sharm/SouthMed/Alex","f59f53d246a0"],"006PW00001kXybtYAC":["hesit","Postponed a month; financial reasons","458bf944f14e"],"006PW00001kYc1RYAS":["wait","Will wait next release for 1BR","ce955faf5f12"],"006PW00001kZfwYYAS":["engaged","Meeting Saturday, 3BR 119m","596c092b1842"],"006PW00001kavw1YAA":["unresp","Busy; sent on WhatsApp","807ff262f978"],"006PW00001khDYUYA2":["engaged","Meeting next Monday, 82m","71fe2e0349ed"],"006PW00001khFX4YAM":["engaged","Meeting Saturday, 3BR/townhouse","9460c1a98f82"],"006PW00001khLarYAE":["wait","Will wait for Alex project","2a9452407847"],"006PW00001khNUeYAM":["engaged","Coming in an hour","b3027288ffc8"],"006PW00001khQNhYAM":["unresp","Planned meeting; no answer","9824b1fb74fc"],"006PW00001khUizYAE":["wait","Waiting Alex release","cd3d1836800a"],"006PW00001khWuVYAU":["engaged","Met with husband; following up","27070e4a445e"],"006PW00001khXgrYAE":["hesit","Broker for Saudi client; two weeks","c688f9e7a082"],"006PW00001khZifYAE":["engaged","May come Saturday; 98m Noor","4ea9ea4fb066"],"006PW00001khcWrYAI":["none","Only requirement: 3BR","1b45e4c62590"],"006PW00001khg4EYAQ":["hesit","Coming in January to get residency/cheques","b5767d51f7a9"],"006PW00001khhczYAA":["wait","Waiting for Alex","ac7eff5e1205"],"006PW00001khjjdYAA":["wait","Son refused Noor; waiting Alex","e24d23cfd481"],"006PW00001khkuDYAQ":["hesit","Not ready; maybe next year","5f4338edf3d5"],"006PW00001khlqHYAQ":["wait","Interested in Alex, depends on budget","27b7eeb676a6"],"006PW00001khmKvYAI":["wait","Prefers Abu Qir; waiting details","814f7a14d0e2"],"006PW00001khpqgYAA":["none","No notes","da39a3ee5e6b"],"006PW00001ki1gXYAQ":["engaged","Zoom today; Privado 2BR RTM","ad62b7080774"],"006PW00001kiBrDYAU":["wait","Waiting Alex launch","b78ab6fd4806"],"006PW00001kiFUzYAM":["engaged","Confirming Sunday meeting, villa","a575ea51ef16"],"006PW00001kiLddYAE":["wait","Waiting for Alex","293722bc69de"],"006PW00001kiNynYAE":["wait","Waiting Alex launch","c60609db4ac1"],"006PW00001kiP69YAE":["wait","Interested in Alexandria","30856c272303"],"006PW00001kiPu9YAE":["wait","Waiting Alex launch","ffe3096d1baa"],"006PW00001kiWFiYAM":["wait","Waiting Alex launch","b78ab6fd4806"],"006PW00001kiYXdYAM":["wait","Wants Alex details when out","869fabb8e570"],"006PW00001kieUzYAI":["engaged","Wants 98m 4th floor; visit to be set","7da009edf942"],"006PW00001kiuBbYAI":["wait","Interested in Alexandria","68f3238ee4dc"],"006PW00001kj3YDYAY":["engaged","Office meeting Sunday","cc7019b7b4d1"],"006PW00001kj4cLYAQ":["wait","Will close 82m at release","df103dd7575b"],"006PW00001kjICIYA2":["wait","Interested in Alexandria","a8ec17b632b6"],"006PW00001kjIqbYAE":["none","Only requirement: 67m Noor","18ef7d4ededd"],"006PW00001kjOZSYA2":["hesit","Will buy after returning in a month","05686cbd1001"],"006PW00001kjZkzYAE":["wait","Wants large RTM ground garden","41c585ef5fe4"],"006PW00001kjcFqYAI":["hesit","Torn between Noor and waiting Alex","bfb135f563bb"],"006PW00001kjue9YAA":["none","Only requirement: 3BR Noor","15330e082bfa"],"006PW00001kk1niYAA":["wait","Waiting Alex release","ae554d30ff9c"],"006PW00001kk5XxYAI":["engaged","Coming Sunday; wants nearer delivery","d607e21a07b5"],"006PW00001kkPGXYA2":["book","Reserved CH3/024/0205; 75K paid","2fd4d45dec28"],"006PW00001klXNJYA2":["wait","Wants Alex airport-land project","3e7270967f52"],"006PW00001kltKdYAI":["wait","Interested in Alexandria","a8ec17b632b6"],"006PW00001kmTKtYAM":["engaged","Discussing direct-marina 2-3BR","3ec301eac9f9"],"006PW00001knOMOYA2":["wait","3BR at Alex launch","2559dec72a14"],"006PW00001koWo5YAE":["wait","Waiting Alex release","5586fa427479"],"006PW00001ku3HyYAI":["engaged","Meeting Wednesday, 98m","2e9b60bf8981"],"006PW00001kuBiYYAU":["engaged","Meeting Monday","e1d1d93a009b"],"006PW00001kuME9YAM":["none","Only requirement: 3BR Noor","223c09fbef15"],"006PW00001kuf5WYAQ":["engaged","Out of Cairo; wants meeting, may buy","2a83dabce380"],"006PW00001kugxeYAA":["wait","Interested in Alex","11816f8b7ae8"],"006PW00001kuz7RYAQ":["engaged","Setting meeting, 82m","b9530519774d"],"006PW00001kvGkzYAE":["engaged","Site visit tomorrow, 67m studio","410d8c2a887a"],"006PW00001kvJxNYAU":["engaged","Setting Noor visit","fa8fd0011191"],"006PW00001kwgn3YAA":["wait","Wants 2BR in Alex","ea1a5b993ed5"],"006PW00001kwh82YAA":["engaged","Deciding with son; can close online","dd51a7a301e8"],"006PW00001kyChdYAE":["wait","Wants 2BR in Alex","743b7bea2d3f"],"006PW00001kz3duYAA":["engaged","Privado meeting tomorrow","9fad1361b17b"],"006PW00001l617ZYAQ":["engaged","Saudi-based; coming in 2 weeks","895a3ac40771"],"006PW00001l668oYAA":["engaged","Site meeting today, 119m","0f5a61a8d87c"],"006PW00001l6UW5YAM":["wait","Waiting Alex project","70674a2c920f"],"006PW00001l6WcjYAE":["wait","Wants ready small shop","083b41c2471d"],"006PW00001l6Y3SYAU":["wait","Waiting new units for 140m typical","fc5a6538d1b4"],"006PW00001l6Y6fYAE":["hesit","Saw shop; arranging affairs","3ab967e2259a"],"006PW00001l6YTFYA2":["engaged","Wants more details on 131m","51a99e95d0f3"],"006PW00001l6YcvYAE":["hesit","Rooms small; decides end of week","c31c2d482ec4"],"006PW00001l6ZYzYAM":["engaged","Zoom 2:30; wants nearer delivery","0c356a7a5822"],"006PW00001l6fL3YAI":["wait","Waiting Alex info; open to Cairo offers","0dabfa241233"],"006PW00001l6ji2YAA":["wait","Interested in Alex","dc674bdcaf20"],"006PW00001l6k9UYAQ":["wait","Wants Alex projects","03962df90f49"],"006PW00001l6kRBYAY":["engaged","Reviewing 123m; reply Saturday","fd3a99177634"],"006PW00001l6mw1YAA":["unresp","Not answering; wanted Alex","18f61e366c3b"],"006PW00001l6o0AYAQ":["none","Only requirement: 1BR 3-yr delivery","c15dfdfb9f73"],"006PW00001l6pYxYAI":["engaged","Meeting tomorrow, 67m","6706048c36c6"],"006PW00001l6rPRYAY":["wait","Waiting Alex project","1d8270673202"],"006PW00001l6zTRYAY":["wait","Asking Alex commercial","6b6c32e766ea"],"006PW00001l7406YAA":["wait","May take 2-3 Alex units","959fdc90cb09"],"006PW00001l7KPrYAM":["none","Only requirement: 3BR","d66095e7cdbb"],"006PW00001l7LYnYAM":["unresp","Meeting unconfirmed; not answering","0cde948c0527"],"006PW00001l7TWLYA2":["engaged","Reviewing 120m; site visit next","ace9e8dd6ec1"],"006PW00001l7WE7YAM":["engaged","Reviewing garden 3BR details","1e312d05ac93"],"006PW00001l7XuyYAE":["unresp","Unavailable; asked 2BR Noor","adc520b7f22a"],"006PW00001l7hKnYAI":["engaged","Meeting this week","88c12e5d60d9"],"006PW00001l7zZTYAY":["wait","Wants Alex details","7557c53773b1"],"006PW00001l8DsXYAU":["engaged","Coming to Madinaty soon; buying in 3 weeks","89ecafda94e7"],"006PW00001l8kZxYAI":["hesit","Gathering down payment; awaiting advisers","94904cc74981"],"006PW00001l8xKHYAY":["wait","Wants Alex Nozha apartment","cad7fcfa2496"],"006PW00001l99uwYAA":["engaged","Wants Marina A/B 122m","5130612efbf0"],"006PW00001lAQldYAG":["hesit","Busy in lecture; send WhatsApp","ec5bf63f8cea"],"006PW00001lBRHTYA4":["engaged","Office meeting tomorrow; liked V25 villa","d8a695abce25"],"006PW00001lBdaVYAS":["engaged","Office meeting tomorrow; liked V25 villa","644a0a38a25e"],"006PW00001lBdiQYAS":["engaged","Meeting tomorrow; liked V25 villa","b2ad62e6dbe6"],"006PW00001lCLt8YAG":["wait","Waiting Alex launch","604f1d8b4958"],"006PW00001lCN94YAG":["hesit","Comparing SouthMed with other projects","6a8457e50589"],"006PW00001lCd6wYAC":["wait","Waiting Alex project","85642ff7088c"],"006PW00001lCfWrYAK":["none","From Alex; only asked 1-2BR Noor","ca4cd5346668"],"006PW00001lCiePYAS":["wait","Wants Alex apartments","b7c0131883d3"],"006PW00001lCwuIYAS":["wait","Alex broker/buyer with interested clients","66e3f9df139b"],"006PW00001lD6VsYAK":["engaged","Coming again tomorrow, Lagoon 2BR","2449c152a29f"],"006PW00001lDJPRYA4":["hesit","Prices high; comparing resale","c7ac8eef0ea0"],"006PW00001lDPWTYA4":["wait","Waiting Alex; pitched Sharm 74m","38ed59b1c172"],"006PW00001lDZVpYAO":["hesit","Comparing resale quattro","529ff33e4cca"],"006PW00001lDfeTYAS":["wait","Bought resale Privado; waiting Alex","a384ae89699d"],"006PW00001lDfeVYAS":["engaged","Coming Thursday or Sunday","5bb74762d089"],"006PW00001lDgcsYAC":["unresp","Thinking with family; not answering","ff32d8609491"],"006PW00001lDjWnYAK":["wait","Wants Alex townhouse","e8ef29c17fa7"],"006PW00001lDjwcYAC":["engaged","Meeting in Zayed tomorrow","3c842037d064"],"006PW00001lDjzqYAC":["none","No notes","da39a3ee5e6b"],"006PW00001lDk6HYAS":["hesit","Bereavement; follow up later","23f5348286b5"],"006PW00001lDk9VYAS":["engaged","Madinaty resident; will visit","11dece00b800"],"006PW00001lDmfxYAC":["wait","Wants 2BR in Alex","4272c4f40db3"],"006PW00001lDn7NYAS":["wait","Wants small Spine shop","f4ffb06e15ea"],"006PW00001lDnc2YAC":["engaged","Contractor; returning to view villa","21d7c8cab356"],"006PW00001lDokzYAC":["none","Only requirement: Privado","94aa4186e7df"],"006PW00001lDp7ZYAS":["hesit","Back abroad; may visit Dec/Jan","cb1ace03a2b1"],"006PW00001lDtccYAC":["wait","Waiting Alex launch","32e81188d986"],"006PW00001lDubtYAC":["wait","Wants 2BR in Alex","c24672383a99"],"006PW00001lE1yMYAS":["wait","Call when Alex launches","cb561157a2ae"],"006PW00001lE32TYAS":["none","Only requirement: 146m","5fbad8862267"],"006PW00001lE55tYAC":["engaged","Meeting Sunday in Zayed","feda7784bfae"],"006PW00001lE6d3YAC":["engaged","RTM 2BR; talk next week","962431d97781"],"006PW00001lE88bYAC":["wait","Wants 2BR in Alex","2afe7e9f0ac0"],"006PW00001lE99VYAS":["wait","Waiting Alex release","5fa3cff98955"],"006PW00001lE9vtYAC":["engaged","Meeting Saturday, 3BR Noor","2afe17d1c7a8"],"006PW00001lEBUgYAO":["wait","Wants Alex details when out","eded38058578"],"006PW00001lEE9BYAW":["engaged","Meeting next Monday; consulting wife","01aed3655a88"],"006PW00001lEFA7YAO":["book","Coming tomorrow; may book Privado 2BR","6b6c7d2df209"],"006PW00001lEGVxYAO":["hesit","Undecided; maybe in 2 months","292b255dfa44"],"006PW00001lEGVyYAO":["hesit","Not before 3 months (new job)","056df83a5ab6"],"006PW00001lEHF7YAO":["engaged","Site visit, Privado RTM 2BR","c189110cf88c"],"006PW00001lEICqYAO":["wait","Interested in Alex","cd51ea3a9467"],"006PW00001lEKD0YAO":["wait","Wants Alex details when out","8568c61a3e6e"],"006PW00001lEKEcYAO":["hesit","Needs a month; busy and travelling","51f4e18d1384"],"006PW00001lEKMgYAO":["wait","Wants small shop in better location","7b944bd5cc32"],"006PW00001lEMgEYAW":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEMjRYAW":["engaged","Interested 82m; payment sent","c0e18e71053e"],"006PW00001lEPZHYA4":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEQGqYAO":["wait","Wants Alex project","b1669e9d6936"],"006PW00001lESCBYA4":["wait","Prefers airport-land project; else Noor","53aa24c25c98"],"006PW00001lET9pYAG":["hesit","Funds in stock market slump; can't act","d11e42c60dd7"],"006PW00001lEVZqYAO":["hesit","Still interested; ready in 2 weeks","7031114ef87e"],"006PW00001lEaknYAC":["engaged","Meeting tomorrow 6pm, Celia 3BR","f0bd9acb9411"],"006PW00001lEcg9YAC":["wait","Waiting Noor 82m release","b6ccc324e664"],"006PW00001lEfz0YAC":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEg8fYAC":["unresp","Difficult; WhatsApp only, not answering","bf18bb2dfeea"],"006PW00001lEhRMYA0":["book","May choose two units tomorrow","77dbc156cc47"],"006PW00001lEiC8YAK":["hesit","Zoom done; praying on it before release","2613f986f96f"],"006PW00001lEiNNYA0":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEjL3YAK":["engaged","Office meeting next week","303bbed3c73b"],"006PW00001lEjxlYAC":["wait","Waiting Alex launch","33c683a74fcc"],"006PW00001lElMrYAK":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEmh7YAC":["engaged","Following Four Seasons details","7de76dc851ac"],"006PW00001lEnLRYA0":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEqpZYAS":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lEsGHYA0":["wait","Waiting Alex launch","0e21efaffc08"],"006PW00001lF1ufYAC":["engaged","Contractor; meeting Wednesday","85806237346f"],"006PW00001lF9U1YAK":["engaged","Referral; possible investment","9938a29af9fc"],"006PW00001lFAhpYAG":["hesit","Maybe early next year","306bba6fd9ab"],"006PW00001lFBpBYAW":["none","Details sent only","aa4dd39eabfc"],"006PW00001lFHg5YAG":["hesit","Visiting site 10 Nov first","79c309694943"],"006PW00001lFPaRYAW":["engaged","Upgrade request; awaiting reply","f3e6e5fd2805"],"006fI000005S56pQAC":["engaged","Zoom planned; asking about Sahel","7d2cf3c145ef"],"006fI0000060iFJQAY":["engaged","Discussing 30m East Hub shop","0ab5ac73e0bf"]}''')


# ============================ 4. MAIN ============================
def main():
    today = date.fromisoformat(REPORT_DATE) if REPORT_DATE else datetime.now(TZ).date()
    # snapshot mode: python pipeline_health.py --users users.json --opps opps.json [--fragment out]
    if "--opps" in sys.argv:
        a = sys.argv
        ou = json.load(open(a[a.index("--opps") + 1]))
        opps = ou["records"] if isinstance(ou, dict) else ou
        users = {}
        for uid, nm, mgr, team in json.load(open(a[a.index("--users") + 1])):
            users[uid] = {"name": nm, "mgr": mgr, "active": True, "team": team or ""}
        for o in opps:      # reps come from the opp owner fields
            ow = o.get("Owner") or {}
            users.setdefault(o["OwnerId"], {"name": ow.get("Name", ""), "mgr": None,
                                            "active": ow.get("IsActive", True),
                                            "team": ow.get("Team_for_Tracking_membership__c") or ""})
        stages = OPEN_STAGES or ["Qualified First Meeting", "Qualified - Waiting Release",
                                 "Qualified Negotiating", "Qualified Closing"]
        base = SF_BASE_URL or ""
    else:
        conn = globals().get("sf")             # the simple_salesforce connection from your notebook
        if conn is None:
            try:
                import __main__; conn = getattr(__main__, "sf", None)
            except Exception: conn = None
        if conn is None and os.environ.get("SF_USERNAME"):
            from simple_salesforce import Salesforce
            env = {k: (os.environ.get(k) or "").strip() for k in
                   ("SF_USERNAME", "SF_PASSWORD", "SF_SECURITY_TOKEN", "SF_DOMAIN")}
            # show WHICH settings arrived (never their values) so a missing GitHub secret is obvious
            print("🔎 Salesforce settings: " + " · ".join(
                f"{k}={'set (' + str(len(v)) + ' chars)' if v else 'MISSING'}" for k, v in env.items()))
            if env["SF_USERNAME"]: print(f"   username: {env['SF_USERNAME']}")
            if not env["SF_SECURITY_TOKEN"]:
                print("   ⚠️ SF_SECURITY_TOKEN is empty — add it under Settings → Secrets and variables → Actions "
                      "→ Repository secrets (not Environment secrets)")
            conn = Salesforce(username=env["SF_USERNAME"], password=env["SF_PASSWORD"],
                              security_token=env["SF_SECURITY_TOKEN"], domain=env["SF_DOMAIN"] or "login")
            print("🔑 connected to Salesforce from SF_USERNAME / SF_PASSWORD / SF_SECURITY_TOKEN")
        if conn is None:
            print("❌ no Salesforce connection: create `sf = Salesforce(...)` first (Colab), "
                  "or set SF_USERNAME, SF_PASSWORD and SF_SECURITY_TOKEN (GitHub / server)."); sys.exit(1)
        print("⏱ pulling users, open stages and open opportunities…")
        users, opps, stages, base = fetch(conn)
    print(f"   {len(opps):,} open opportunities · stages: {stages}")
    global NOTE_LABELS_FILE
    if "--labels" in sys.argv:
        NOTE_LABELS_FILE = sys.argv[sys.argv.index("--labels") + 1]
    if "--merge-labels" in sys.argv:
        # CSV columns: Opportunity ID, Stance (book/engaged/wait/hesit/unresp/lost/none), Summary
        import csv
        valid = {k for k, _l, _h in STANCES} | {NO_NOTES[0]}
        by_id = {o["Id"]: o for o in opps}
        store = {}
        if os.path.exists(NOTE_LABELS_FILE):
            with open(NOTE_LABELS_FILE, encoding="utf-8") as f: store = json.load(f)
        added, bad = 0, []
        with open(sys.argv[sys.argv.index("--merge-labels") + 1], encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                oid, st = (r.get("Opportunity ID") or "").strip(), (r.get("Stance") or "").strip()
                if oid not in by_id or st not in valid:
                    bad.append(oid or "?"); continue
                raw = "\n".join(str(by_id[oid].get(x) or "") for x in NOTE_FIELDS if by_id[oid].get(x))
                store[oid] = {"stance": st, "summary": (r.get("Summary") or "").strip(),
                              "note_hash": note_hash(clean_note(raw)), "reviewed": today.isoformat()}
                added += 1
        with open(NOTE_LABELS_FILE, "w", encoding="utf-8") as f: json.dump(store, f, ensure_ascii=False, indent=0)
        print(f"   🧾 merged {added:,} reviewed labels into {NOTE_LABELS_FILE}" + (f" · skipped {len(bad)}: {bad[:10]}" if bad else ""))
    labels = load_labels(use_sheet="--opps" not in sys.argv and "--no-sheet" not in sys.argv)
    to_review = []
    rows, meta = build(users, opps, stages, today, base, labels=labels, to_review=to_review)
    if "--export-labels" in sys.argv:
        export_labels(opps, labels, sys.argv[sys.argv.index("--export-labels") + 1])
    if ("--opps" not in sys.argv and "--no-sheet" not in sys.argv) or "--write-all-notes" in sys.argv:
        write_all_notes(opps, rows, meta)
    if to_review:
        import csv
        rv_out = f"Notes_To_Review_{today:%Y%m%d}.csv"
        with open(rv_out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(to_review[0])); w.writeheader(); w.writerows(to_review)
        print(f"   📝 {len(to_review):,} new / changed notes not reviewed yet → {os.path.abspath(rv_out)}")
    else:
        print("   ✅ every note has a reviewed stance")
    frag = "--fragment" in sys.argv
    out = (sys.argv[sys.argv.index("--fragment") + 1] if frag else None) or \
          f"{FILE_TAG}_{today.day}{'th' if 10 <= today.day % 100 <= 20 else {1:'st',2:'nd',3:'rd'}.get(today.day % 10,'th')}{today:%b}.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(render(rows, meta, fragment=frag))
    h = [sum(1 for r in rows if r["h"] == i) for i in range(3)]
    print(f"   healthy {h[0]:,} · watch {h[1]:,} · at risk {h[2]:,} · value {sum(r['amt'] for r in rows):,.0f} {CURRENCY}")
    print(f"📄 {os.path.abspath(out)}")


main()
