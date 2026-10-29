# -*- coding: utf-8 -*-
"""
zotero_link_citations.py
========================
Word(.docx) 원고 본문에 텍스트로 적힌 저자-연도 인용, 예: (Yuk et al., 2022)
를 내 Zotero 라이브러리의 '특정 폴더(컬렉션)' 항목과 매칭하여
Zotero 활성 필드(ZOTERO_ITEM)로 바꿔 줍니다.
문서 끝의 텍스트 참고문헌 목록은 Zotero 참고문헌 필드(ZOTERO_BIBL)로 바꾸고,
문서 스타일은 Human Reproduction으로 미리 설정합니다.

동작 순서
  1) zotero.sqlite 를 임시 폴더로 복사해서 읽기 전용으로 엽니다 (원본 DB는 건드리지 않음).
  2) 검색 범위 폴더 선택: Zotero에서 지금 선택해 둔 폴더를 자동 감지 → 확인.
     (아니면 목록에서 고르기, 또는 --collection "폴더명")
  3) 문서 끝 References 목록의 항목마다 폴더 안의 Zotero 항목을 찾습니다
     (DOI → 제1저자+연도 → 제목 유사도 순서).
     확실하지 않거나 못 찾으면 사용자에게 물어봅니다.
  4) 본문 인용 (저자 et al., 연도)를 References 항목과 이어서 Zotero 항목으로 연결하고,
     해당 텍스트를 Zotero 필드 코드로 감쌉니다 (글꼴과 기울임 등 서식은 그대로 유지).
  5) 결과는 <원본이름>_zotero.docx 로 저장합니다 (원본은 그대로 둠).
     Word에서 열고 Zotero 탭 → Refresh 를 누르면 Human Reproduction 스타일로 다시 그려집니다.

필요 패키지 (Anaconda Prompt):
    pip install python-docx lxml rapidfuzz

사용 예:
    python zotero_link_citations.py "D:\\원고\\manuscript.docx"
    python zotero_link_citations.py manuscript.docx --collection "LNG-IUS osteoporosis"
    python zotero_link_citations.py            (파일 선택 창이 뜹니다)
"""

import argparse
import copy
import json
import os
import random
import re
import shutil
import sqlite3
import string
import sys
import tempfile
import unicodedata
import urllib.request
import zipfile

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

try:
    from docx import Document
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("python-docx 가 필요합니다:  pip install python-docx")

try:
    from rapidfuzz import fuzz as _fuzz

    def partial_ratio(a, b):
        return _fuzz.partial_ratio(a, b)

    def token_ratio(a, b):
        return _fuzz.token_set_ratio(a, b)
except ImportError:  # rapidfuzz 가 없으면 difflib 로 대체 (조금 느림)
    import difflib

    def partial_ratio(a, b):
        if not a or not b:
            return 0
        short, long_ = (a, b) if len(a) <= len(b) else (b, a)
        sm = difflib.SequenceMatcher(None, short, long_, autojunk=False)
        best = 0
        for blk in sm.get_matching_blocks():
            start = max(0, blk.b - blk.a)
            sub = long_[start:start + len(short)]
            r = difflib.SequenceMatcher(None, short, sub, autojunk=False).ratio()
            best = max(best, r)
            if best == 1:
                break
        return best * 100

    def token_ratio(a, b):
        return difflib.SequenceMatcher(None, " ".join(sorted(a.split())),
                                       " ".join(sorted(b.split()))).ratio() * 100


# ------------------------------------------------------------------ 기본 설정
DEFAULT_ZOTERO_DIR = r"D:\OneDrive\zotero"
STYLE_ID = "http://www.zotero.org/styles/human-reproduction"
CONNECTOR_URL = "http://127.0.0.1:23119/connector/getSelectedCollection"
REF_HEADINGS = {"references", "reference", "bibliography", "literature cited",
                "참고문헌", "참고 문헌"}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# ------------------------------------------------------------------ 문자열 도우미
def norm(s):
    """악센트 제거, 소문자, 영숫자와 공백만 남김."""
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("’", "'")
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"_", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_name(s):
    """성(姓) 비교용: 공백, 하이픈, 아포스트로피 모두 제거."""
    return re.sub(r"[\s\-']", "", norm(s))


def short(s, n=95):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[: n - 1] + "…"


def ask(prompt):
    try:
        return input(prompt)
    except EOFError:
        return ""


def rand_id(n=8):
    return "".join(random.choice(string.ascii_letters + string.digits) for _ in range(n))


# ================================================================== Zotero DB
class ZoteroDB:
    CSL_TYPES = {
        "journalArticle": "article-journal", "book": "book", "bookSection": "chapter",
        "conferencePaper": "paper-conference", "report": "report", "thesis": "thesis",
        "webpage": "webpage", "magazineArticle": "article-magazine",
        "newspaperArticle": "article-newspaper", "preprint": "article",
        "manuscript": "manuscript", "document": "article", "dataset": "dataset",
        "encyclopediaArticle": "entry-encyclopedia", "dictionaryEntry": "entry-dictionary",
        "patent": "patent", "statute": "legislation", "case": "legal_case",
        "presentation": "speech", "computerProgram": "software", "blogPost": "post-weblog",
        "standard": "standard", "letter": "personal_communication",
    }
    CSL_FIELDS = {
        "title": "title", "publicationTitle": "container-title", "bookTitle": "container-title",
        "proceedingsTitle": "container-title", "websiteTitle": "container-title",
        "encyclopediaTitle": "container-title", "dictionaryTitle": "container-title",
        "journalAbbreviation": "container-title-short", "shortTitle": "title-short",
        "volume": "volume", "issue": "issue", "pages": "page", "DOI": "DOI",
        "ISSN": "ISSN", "ISBN": "ISBN", "url": "URL", "publisher": "publisher",
        "place": "publisher-place", "edition": "edition", "language": "language",
        "series": "collection-title", "seriesNumber": "collection-number",
        "numberOfVolumes": "number-of-volumes", "numPages": "number-of-pages",
        "institution": "publisher", "university": "publisher", "reportNumber": "number",
        "reportType": "genre", "thesisType": "genre", "archiveID": "number",
        "repository": "publisher", "abstractNote": "abstract",
    }
    CSL_CREATORS = {"author": "author", "editor": "editor", "bookAuthor": "container-author",
                    "seriesEditor": "collection-editor", "translator": "translator",
                    "director": "author", "inventor": "author", "presenter": "author",
                    "programmer": "author", "artist": "author", "interviewee": "author",
                    "contributor": None, "reviewedAuthor": None}

    def __init__(self, zotero_dir):
        src = os.path.join(zotero_dir, "zotero.sqlite")
        if not os.path.exists(src):
            sys.exit(f"[오류] {src} 가 없습니다. --zotero-dir 로 Zotero 데이터 폴더를 지정하세요.")
        self.tmpdir = tempfile.mkdtemp(prefix="zotlink_")
        dst = os.path.join(self.tmpdir, "zotero.sqlite")
        try:
            shutil.copy2(src, dst)
        except (PermissionError, OSError):
            bak = src + ".bak"
            if not os.path.exists(bak):
                sys.exit("[오류] Zotero DB를 복사할 수 없습니다 (잠겨 있음). Zotero를 잠시 닫고 다시 실행하세요.")
            print("[주의] zotero.sqlite 가 잠겨 있어 zotero.sqlite.bak (마지막 백업)을 사용합니다.")
            shutil.copy2(bak, dst)
        self.con = sqlite3.connect(dst)
        self.con.row_factory = sqlite3.Row
        self._load()

    def close(self):
        try:
            self.con.close()
            shutil.rmtree(self.tmpdir, ignore_errors=True)
        except Exception:
            pass

    def _has(self, name):
        return self.con.execute("SELECT 1 FROM sqlite_master WHERE name=?", (name,)).fetchone() is not None

    def q(self, sql, args=()):
        return self.con.execute(sql, args).fetchall()

    def _load(self):
        it = "itemTypesCombined" if self._has("itemTypesCombined") else "itemTypes"
        fl = "fieldsCombined" if self._has("fieldsCombined") else "fields"
        deleted = "AND i.itemID NOT IN (SELECT itemID FROM deletedItems)" if self._has("deletedItems") else ""

        # 라이브러리 URI 정보
        settings = {}
        if self._has("settings"):
            for r in self.q("SELECT setting, key, value FROM settings"):
                settings[(r["setting"], r["key"])] = r["value"]
        self.user_id = settings.get(("account", "userID"))
        self.local_key = settings.get(("account", "localUserKey"))
        self.lib_type, self.group_id = {}, {}
        if self._has("libraries"):
            for r in self.q("SELECT libraryID, type FROM libraries"):
                self.lib_type[r["libraryID"]] = r["type"]
        if self._has("groups"):
            for r in self.q("SELECT groupID, libraryID FROM groups"):
                self.group_id[r["libraryID"]] = r["groupID"]

        # 항목 (첨부, 노트, 주석, 휴지통 제외)
        self.items = {}
        for r in self.q(f"""SELECT i.itemID, i.key, i.libraryID, t.typeName FROM items i
                            JOIN {it} t ON i.itemTypeID = t.itemTypeID
                            WHERE t.typeName NOT IN ('attachment','note','annotation') {deleted}"""):
            self.items[r["itemID"]] = {"id": r["itemID"], "key": r["key"], "lib": r["libraryID"],
                                       "type": r["typeName"], "fields": {}, "creators": []}
        for r in self.q(f"""SELECT d.itemID, f.fieldName, v.value FROM itemData d
                            JOIN {fl} f ON d.fieldID = f.fieldID
                            JOIN itemDataValues v ON d.valueID = v.valueID"""):
            if r["itemID"] in self.items:
                self.items[r["itemID"]]["fields"][r["fieldName"]] = str(r["value"])
        for r in self.q("""SELECT ic.itemID, c.firstName, c.lastName, c.fieldMode, ct.creatorType
                           FROM itemCreators ic JOIN creators c ON ic.creatorID = c.creatorID
                           JOIN creatorTypes ct ON ic.creatorTypeID = ct.creatorTypeID
                           ORDER BY ic.itemID, ic.orderIndex"""):
            if r["itemID"] in self.items:
                self.items[r["itemID"]]["creators"].append(
                    {"first": r["firstName"] or "", "last": r["lastName"] or "",
                     "single": r["fieldMode"] == 1, "type": r["creatorType"]})
        for itm in self.items.values():
            f = itm["fields"]
            m = re.match(r"(\d{4})", f.get("date", ""))
            itm["year"] = m.group(1) if m else ""
            auth = [c for c in itm["creators"] if c["type"] == "author"] or itm["creators"]
            itm["authors"] = [c["last"] for c in auth]
            itm["title"] = f.get("title", "")
            itm["doi"] = f.get("DOI", "").lower().strip()
            if not itm["doi"]:
                m = re.search(r"\b(10\.\d{4,9}/\S+)", f.get("extra", ""))
                itm["doi"] = m.group(1).lower().rstrip(".") if m else ""

        # 컬렉션
        self.collections = {}
        for r in self.q("SELECT collectionID, collectionName, parentCollectionID, libraryID FROM collections"):
            self.collections[r["collectionID"]] = {"id": r["collectionID"], "name": r["collectionName"],
                                                   "parent": r["parentCollectionID"], "lib": r["libraryID"]}
        self.coll_items = {}
        for r in self.q("SELECT collectionID, itemID FROM collectionItems"):
            self.coll_items.setdefault(r["collectionID"], set()).add(r["itemID"])

    # --- 컬렉션 도우미
    def coll_path(self, cid):
        parts, c = [], self.collections.get(cid)
        while c:
            parts.append(c["name"])
            c = self.collections.get(c["parent"])
        return " / ".join(reversed(parts))

    def descendants(self, cid):
        out, stack = [cid], [cid]
        while stack:
            cur = stack.pop()
            for c in self.collections.values():
                if c["parent"] == cur:
                    out.append(c["id"])
                    stack.append(c["id"])
        return out

    def items_in(self, cids):
        ids = set()
        for c in cids:
            ids |= self.coll_items.get(c, set())
        return [self.items[i] for i in ids if i in self.items]

    # --- Word 필드에 넣을 정보
    def uris(self, itm):
        lib = itm["lib"]
        if self.lib_type.get(lib) == "group" and lib in self.group_id:
            return [f"http://zotero.org/groups/{self.group_id[lib]}/items/{itm['key']}"]
        out = []
        if self.user_id:
            out.append(f"http://zotero.org/users/{self.user_id}/items/{itm['key']}")
        if self.local_key:
            out.append(f"http://zotero.org/users/local/{self.local_key}/items/{itm['key']}")
        return out

    def csl(self, itm):
        f = itm["fields"]
        d = {"id": itm["id"], "type": self.CSL_TYPES.get(itm["type"], "article")}
        for zf, cf in self.CSL_FIELDS.items():
            if f.get(zf) and cf not in d:
                d[cf] = f[zf]
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", f.get("date", ""))
        if m:
            y, mo, dd = (int(x) for x in m.groups())
            parts = [y] + ([mo] if mo else []) + ([dd] if mo and dd else [])
            d["issued"] = {"date-parts": [parts]}
        for c in itm["creators"]:
            key = self.CSL_CREATORS.get(c["type"], "author")
            if not key:
                continue
            name = {"literal": c["last"]} if c["single"] else {"family": c["last"], "given": c["first"]}
            d.setdefault(key, []).append(name)
        return d

    def label(self, itm):
        a = itm["authors"]
        au = (a[0] if a else "?") + (" et al." if len(a) > 2 else (f" & {a[1]}" if len(a) == 2 else ""))
        jr = itm["fields"].get("journalAbbreviation") or itm["fields"].get("publicationTitle", "")
        return f"{au} ({itm['year'] or 'n.d.'}) {short(itm['title'], 70)}" + (f" — {short(jr, 25)}" if jr else "")


# ================================================================== 폴더 선택
def detect_selected_collection(db):
    """Zotero가 실행 중이면 지금 왼쪽 창에서 선택된 폴더를 알려 줌 (Connector API)."""
    try:
        req = urllib.request.Request(CONNECTOR_URL, data=b"{}", method="POST",
                                     headers={"Content-Type": "application/json",
                                              "X-Zotero-Connector-API-Version": "3"})
        with urllib.request.urlopen(req, timeout=3) as r:
            data = json.loads(r.read().decode("utf-8"))
    except Exception:
        return None
    cid = data.get("id")
    if isinstance(cid, str):
        cid = int(cid[1:]) if cid[:1] == "C" and cid[1:].isdigit() else (int(cid) if cid.isdigit() else None)
    if cid in db.collections:
        return cid
    name = data.get("name")  # id가 안 맞으면 이름으로 찾기
    hits = [c["id"] for c in db.collections.values()
            if c["name"] == name and c["lib"] == data.get("libraryID", c["lib"])]
    return hits[0] if len(hits) == 1 else None


def choose_collection_from_list(db):
    rows = sorted(db.collections, key=lambda c: db.coll_path(c).lower())
    flt = ""
    while True:
        shown = [c for c in rows if flt.lower() in db.coll_path(c).lower()]
        print("\n  Zotero 폴더 목록" + (f" (필터: '{flt}')" if flt else ""))
        for n, c in enumerate(shown[:60], 1):
            print(f"   {n:3d}) {db.coll_path(c)}  [{len(db.items_in([c]))}]")
        if len(shown) > 60:
            print(f"   ... {len(shown) - 60}개 더 있음 — 이름 일부를 입력해 좁혀 보세요")
        a = ask("  번호 선택 / 이름 일부 입력(필터) / Enter=필터 해제 > ").strip()
        if a.isdigit() and 1 <= int(a) <= min(60, len(shown)):
            return shown[int(a) - 1]
        flt = a


def pick_scope(db, args):
    cid = None
    if args.collection:
        hits = [c for c in db.collections if db.collections[c]["name"].lower() == args.collection.lower()
                or db.coll_path(c).lower() == args.collection.lower()]
        if len(hits) == 1:
            cid = hits[0]
        else:
            print(f"[주의] '{args.collection}' 폴더를 하나로 특정할 수 없습니다 ({len(hits)}개). 목록에서 선택하세요.")
            cid = choose_collection_from_list(db)
    else:
        cid = detect_selected_collection(db)
        if cid is not None:
            print(f"\nZotero에서 지금 선택된 폴더: '{db.coll_path(cid)}' (항목 {len(db.items_in([cid]))}개)")
            a = "y" if args.yes else ask("  이 폴더에서만 찾을까요? [Y=예 / n=다른 폴더 선택] > ").strip().lower()
            if a.startswith("n"):
                cid = choose_collection_from_list(db)
        else:
            print("\n(Zotero가 꺼져 있거나 선택된 폴더를 알 수 없어 목록에서 고릅니다.)")
            cid = choose_collection_from_list(db)
    cids = [cid]
    subs = db.descendants(cid)
    if len(subs) > 1:
        if args.yes:
            a = "y"
        else:
            a = ask(f"  하위 폴더 {len(subs) - 1}개도 포함할까요? [Y/n] > ").strip().lower()
        if not a.startswith("n"):
            cids = subs
    items = db.items_in(cids)
    print(f"→ 검색 범위: '{db.coll_path(cid)}'" + (" (+하위 폴더)" if len(cids) > 1 else "") +
          f", 항목 {len(items)}개\n")
    return cid, items


# ================================================================== 참고문헌 목록 파싱
YEAR_RE = re.compile(r"\b((?:19|20)\d{2})([a-z]?)\b")
DOI_RE = re.compile(r"\b(10\.\d{4,9}/[^\s\"<>]+)", re.I)
PARTICLES = r"(?:(?:van|von|de|der|den|du|da|di|la|le|del|dos|des|el|al|ter|ten)\s+)*"


class Ref:
    def __init__(self, idx, text, p_el):
        self.idx, self.text, self.p = idx, text.strip(), p_el
        self.item = None
        self.status = ""
        m = DOI_RE.search(self.text)
        self.doi = m.group(1).rstrip(".,;").lower() if m else ""
        # 저자 부분: 첫 ". " 이전 (Vancouver / Human Reproduction 형식)
        auth_block = re.split(r"\.\s+(?=[A-Z0-9\"'“(])", self.text, maxsplit=1)[0]
        if re.search(r"\(\s*(19|20)\d{2}", auth_block):  # APA 형식: Yuk, J. S., ... (2022).
            auth_block = re.split(r"\(\s*(?:19|20)\d{2}", auth_block)[0]
        names = [a.strip() for a in re.split(r",\s*|\s+&\s+|\s+and\s+", auth_block) if a.strip()]
        self.surnames = []
        for nm in names:
            toks = nm.split()
            # 'Yuk JS' → Yuk ; 'Melton LJ 3rd' → Melton ; 'Yuk, J. S.' 의 이니셜 토큰 제거
            toks = [t for t in toks if not re.fullmatch(r"(?:[A-Z]\.?-?){1,4}|\d+(st|nd|rd|th)|Jr\.?|Sr\.?", t)]
            if toks:
                self.surnames.append(" ".join(toks))
        self.n_authors = len(self.surnames) + (99 if re.search(r"\bet al\b", auth_block) else 0)
        years = YEAR_RE.findall(self.text)
        self.years = {y for y, _ in years}
        self.year_sfx = {y + s for y, s in years}
        m = re.search(r"\b((?:19|20)\d{2})[a-z]?\s*[;(:]", self.text) or re.search(r"\b((?:19|20)\d{2})", self.text)
        self.year = m.group(1) if m else ""

    @property
    def first(self):
        return self.surnames[0] if self.surnames else ""


def find_ref_section(doc):
    """References 제목 문단 위치, 참고문헌 문단 목록, 본문 문단 목록 반환."""
    body = doc.element.body
    paras = list(body.iter(W + "p"))
    head_i = None
    for i, p in enumerate(paras):
        txt = "".join(t.text or "" for t in p.iter(W + "t")).strip().rstrip(":").lower()
        if txt in REF_HEADINGS:
            head_i = i  # 마지막 'References' 사용 (목차 등 회피)
    if head_i is None:
        return None, [], paras
    refs = []
    blank = 0
    for p in paras[head_i + 1:]:
        para = Paragraph(p, doc)
        txt = para.text.strip()
        sname = (para.style.name if para.style is not None else "") or ""
        if sname.lower().startswith(("heading", "title")) or re.match(r"(?i)^(table|figure|supplement|appendix)\b", txt):
            break
        if p.getparent().tag != body.tag:  # 표 안의 문단 등
            break
        if not txt:
            blank += 1
            if blank >= 2:
                break
            continue
        blank = 0
        refs.append(Ref(len(refs) + 1, txt, p))
    return paras[head_i], refs, paras[:head_i]


# ================================================================== 매칭 엔진
class Matcher:
    def __init__(self, db, scope_items, args, memory):
        self.db, self.items, self.args, self.mem = db, scope_items, args, memory

    # 참고문헌 한 줄 ↔ Zotero 항목: (저자 일치, 연도 일치, 제목 유사도)
    def facts(self, ref, itm):
        au = bool(itm["authors"] and ref.first and norm_name(itm["authors"][0]) == norm_name(ref.first))
        yr = bool(itm["year"] and itm["year"] in ref.years)
        t = norm(itm["title"])
        tt = partial_ratio(t, norm(ref.text)) if t else 0
        return au, yr, tt

    def score(self, ref, itm):
        if ref.doi and itm["doi"] and ref.doi == itm["doi"]:
            return 200
        s = 0.0
        if itm["authors"] and ref.first and norm_name(itm["authors"][0]) == norm_name(ref.first):
            s += 30
        if itm["year"] and itm["year"] in ref.years:
            s += 20
        t = norm(itm["title"])
        if t:
            pr = partial_ratio(t, norm(ref.text))
            s += pr * (0.5 if len(t) >= 15 else 0.25)  # 아주 짧은 제목은 가중치 낮춤
        return s

    def rank(self, ref, pool):
        return sorted(((self.score(ref, i), i) for i in pool), key=lambda x: -x[0])

    def auto_pick(self, ref, pool):
        """확실할 때만 항목을 돌려줌. 애매하면 None."""
        doi_hits = [i for i in pool if ref.doi and i["doi"] == ref.doi]
        if len(doi_hits) == 1:
            return doi_hits[0]
        # 제1저자 + 연도가 모두 맞는 후보들
        ay = []
        for i in pool:
            au, yr, tt = self.facts(ref, i)
            if au and yr:
                ay.append((tt, i))
        ay.sort(key=lambda x: -x[0])
        if len(ay) == 1 and ay[0][0] >= 70:
            return ay[0][1]  # 유일한 저자+연도 후보이고 제목도 대체로 맞음
        if len(ay) > 1 and ay[0][0] >= 90 and ay[1][0] < 75:
            return ay[0][1]  # 같은 저자·연도 여러 편 → 제목이 확실히 맞는 것
        if not ay:
            ranked = self.rank(ref, pool)
            if ranked:
                au, yr, tt = self.facts(ref, ranked[0][1])
                # 저자 철자 또는 연도가 다를 때: 긴 제목이 거의 완전히 일치하면 자동
                if (au or yr) and tt >= 97 and len(norm(ranked[0][1]["title"])) >= 30:
                    return ranked[0][1]
        return None

    def match_refs(self, refs):
        print(f"참고문헌 {len(refs)}개를 Zotero 폴더 항목과 대조합니다...")
        pending = []
        for ref in refs:  # 1차: 자동 매칭
            key = "ref::" + norm(ref.text)
            if key in self.mem and not self.args.fresh:
                k = self.mem[key]
                ref.item = next((i for i in self.db.items.values() if i["key"] == k), None) if k else None
                if ref.item or not k:
                    ref.status = "이전 선택" if k else "건너뜀(이전 선택)"
                    continue
            ref.item = self.auto_pick(ref, self.items)
            if ref.item:
                ref.status = "자동"
                self.mem[key] = ref.item["key"]
            else:
                pending.append(ref)
        print(f"  자동 연결 {len(refs) - len(pending)}개, 확인 필요 {len(pending)}개")

        if pending and not self.args.auto:
            if len(pending) >= max(5, 0.4 * len(refs)):
                print(f"\n[주의] {len(pending)}개를 이 폴더에서 찾지 못했습니다. 폴더 선택이 맞나요?")
                a = ask("  Enter = 하나씩 확인 | L = 전체 라이브러리에서 자동으로 다시 찾은 뒤 나머지 확인 | k = 모두 건너뛰기 > ").strip().lower()
                if a == "k":
                    for r in pending:
                        self.mem["ref::" + norm(r.text)] = None
                    pending = []
                elif a == "l":
                    everything = list(self.db.items.values())
                    still = []
                    for r in pending:
                        r.item = self.auto_pick(r, everything)
                        if r.item:
                            r.status = "자동(전체 라이브러리)"
                            self.mem["ref::" + norm(r.text)] = r.item["key"]
                            print(f"    #{r.idx} → {self.db.label(r.item)}")
                        else:
                            still.append(r)
                    print(f"  전체 라이브러리에서 {len(pending) - len(still)}개 추가 연결")
                    pending = still
            for n, ref in enumerate(pending, 1):  # 2차: 사용자 확인
                ranked = self.rank(ref, self.items)
                ref.item = self.interactive(
                    f"({n}/{len(pending)}) 참고문헌 #{ref.idx} 를 Zotero 폴더에서 확실히 찾지 못했습니다.",
                    f"참고문헌: {short(ref.text, 160)}", ranked, ref)
                ref.status = "사용자 선택" if ref.item else "건너뜀"
                self.mem["ref::" + norm(ref.text)] = ref.item["key"] if ref.item else None
        ok = sum(1 for r in refs if r.item)
        print(f"  → {ok}/{len(refs)}개 연결됨\n")

    def interactive(self, title, detail, ranked, ref=None, hint=None):
        if self.args.auto:
            return None
        print("-" * 72)
        print("[확인 필요] " + title)
        print("  " + detail)
        cands = [i for s, i in ranked[:6] if s > 25] if ranked else []
        while True:
            if cands:
                print("  후보:")
                for n, itm in enumerate(cands, 1):
                    sc = ""
                    if ref is not None:
                        sc = f"[{min(100, int(self.score(ref, itm) / 1.0)):3d}점] "
                    print(f"    {n}) {sc}{self.db.label(itm)}")
            else:
                print("  (폴더 안에 비슷한 후보가 없습니다)")
            a = ask("  후보 번호 입력(예: 1) | s [검색어] = 폴더 안 검색 | a [검색어] = 전체 라이브러리 검색 | Enter = 건너뛰기(텍스트 유지)\n  > ").strip()
            if not a:
                return None
            if a.isdigit() and 1 <= int(a) <= len(cands):
                return cands[int(a) - 1]
            m = re.match(r"^([sa])(?:\s+(.*))?$", a, re.I)
            if m:
                where = "폴더 안" if m.group(1).lower() == "s" else "전체 라이브러리"
                pool = self.items if m.group(1).lower() == "s" else list(self.db.items.values())
                term = (m.group(2) or "").strip()
                if term:
                    toks = norm(term).split()
                    hits = [i for i in pool if all(t in norm(" ".join([i["title"], " ".join(i["authors"]), i["year"]])) for t in toks)]
                    if ref is not None:
                        hits.sort(key=lambda i: -self.score(ref, i))
                elif ref is not None:
                    # 검색어 없이 a / s 만 입력 → 이 참고문헌과 가장 비슷한 항목을 자동으로 찾음
                    hits = [i for s_, i in self.rank(ref, pool) if s_ > 25]
                elif hint:
                    toks = norm(hint).split()
                    hits = [i for i in pool if all(t in norm(" ".join([" ".join(i["authors"]), i["year"]])) for t in toks)]
                else:
                    print("  검색어를 함께 입력하세요. 예: a Lethaby")
                    continue
                print(f"  ({where} 검색" + (f": '{term}'" if term else ": 이 참고문헌과 비슷한 항목") + ")")
                cands = hits[:9]
                if not cands:
                    print("  검색 결과 없음.")
                continue
            print("  입력을 이해하지 못했습니다. 후보 번호(예: 1), 'a 검색어', 또는 Enter 를 입력하세요.")

    # 본문 인용 하나 → Zotero 항목
    def resolve_cite(self, a1, a2, etal, year, sfx, refs, display):
        key = f"cite::{norm(display)}::{norm(a1)}::{year}{sfx}"
        if key in self.mem and not self.args.fresh:
            k = self.mem[key]
            if not k:
                return None
            hit = next((i for i in self.db.items.values() if i["key"] == k), None)
            if hit:
                return hit
        a1n = norm_name(a1)
        cands = [r for r in refs if norm_name(r.first) == a1n and year in r.years]
        if a2:
            c2 = [r for r in cands if len(r.surnames) >= 2 and norm_name(r.surnames[1]) == norm_name(a2)]
            cands = c2 or cands
        elif etal:
            c3 = [r for r in cands if r.n_authors >= 3]
            cands = c3 or cands
        else:
            c1 = [r for r in cands if r.n_authors == 1]
            cands = c1 or cands
        if sfx:
            withsfx = [r for r in cands if (year + sfx) in r.year_sfx]
            if withsfx:
                cands = withsfx
            elif len(cands) > 1:  # 목록에 a/b 표기가 없으면 목록 순서로 a=첫째
                i = ord(sfx) - ord("a")
                cands = [cands[i]] if i < len(cands) else cands
        item = None
        if len(cands) == 1 and cands[0].item:
            item = cands[0].item
        elif len(cands) == 1 and not cands[0].item:
            item = None  # 참고문헌은 있으나 사용자가 건너뜀
        else:
            # 참고문헌 목록에서 못 찾음/여러 개 → 폴더 안에서 직접 저자+연도 검색
            pool = [i for i in self.items if i["authors"] and norm_name(i["authors"][0]) == a1n and i["year"] == year]
            if cands:  # 여러 참고문헌 후보 중 연결된 항목만
                linked = [r.item for r in cands if r.item]
                pool = linked or pool
            if a2:
                p2 = [i for i in pool if len(i["authors"]) >= 2 and norm_name(i["authors"][1]) == norm_name(a2)]
                pool = p2 or pool
            if len(pool) == 1:
                item = pool[0]
            else:
                ranked = [(100, i) for i in pool]
                if not ranked:  # 성만 같거나 연도만 같은 것까지 보여줌
                    loose = [i for i in self.items if (i["authors"] and norm_name(i["authors"][0]) == a1n) or
                             (i["year"] == year and i["authors"] and partial_ratio(a1n, norm_name(i["authors"][0])) > 80)]
                    ranked = [(50, i) for i in loose]
                item = self.interactive(f"본문 인용 '{display}' 에 해당하는 항목을 확정할 수 없습니다.",
                                        f"찾는 조건: 제1저자 '{a1}', 연도 {year}{sfx}" +
                                        (f" (참고문헌 후보 {len(cands)}개)" if cands else " (참고문헌 목록에서도 못 찾음)"),
                                        ranked, hint=f"{a1} {year}")
        self.mem[key] = item["key"] if item else None
        return item


# ================================================================== 본문 인용 탐지
NAME = r"[^\W\d_][\w'’\-]*"
AUTH = PARTICLES + r"[A-Z][\w'’\-]*(?:\s+(?:of|for|the|on|in)?\s*[A-Z][\w'’\-]*){0,4}?"
YEARS = r"(?P<years>(?:19|20)\d{2}[a-z]?(?:\s*,\s*(?:19|20)\d{2}[a-z]?)*)"
PART_RE = re.compile(
    r"^(?P<prefix>(?:e\.g\.,?|i\.e\.,?|see(?: also)?|cf\.|reviewed (?:in|by)|as reviewed (?:in|by))\s+)?"
    r"(?P<a1>" + AUTH + r")"
    r"(?:\s+(?:and|&)\s+(?P<a2>" + AUTH + r")|(?P<etal>\s+et\s+al\.?))?"
    r",?\s+" + YEARS + r"$")
NARR_RE = re.compile(
    r"(?<![\w(])(?P<a1>" + AUTH + r")(?:\s+(?:and|&)\s+(?P<a2>" + AUTH + r")|(?P<etal>\s+et\s+al\.?))?"
    r"\s+(?P<paren>\(" + YEARS + r"\))")
PAREN_RE = re.compile(r"\(([^()]{4,800})\)")
SKIP_WORDS = {"Table", "Figure", "Fig", "Supplement", "Appendix", "HR", "OR", "CI", "No", "In", "The", "From", "Since"}


def parse_parts(inner):
    out = []
    for part in [p.strip() for p in inner.split(";")]:
        m = PART_RE.match(part)
        if not m or m.group("a1").split()[0] in SKIP_WORDS:
            return None
        ys = [y.strip() for y in m.group("years").split(",")]
        for y in ys:
            out.append({"prefix": (m.group("prefix") or "").strip(), "a1": m.group("a1"), "a2": m.group("a2"),
                        "etal": bool(m.group("etal")), "year": y[:4], "sfx": y[4:]})
    return out


# ================================================================== Word XML 조작
def run_text(r):
    s = ""
    for ch in r:
        if ch.tag == W + "t":
            s += ch.text or ""
        elif ch.tag in (W + "tab",):
            s += "\t"
        elif ch.tag in (W + "br", W + "cr"):
            s += "\n"
        elif ch.tag in (W + "noBreakHyphen",):
            s += "-"
        elif ch.tag in (W + "drawing", W + "object", W + "pict", W + "footnoteReference", W + "endnoteReference"):
            s += "\uFFFC"
    return s


def para_map(p):
    """문단의 run 목록과 이어 붙인 텍스트. 기존 필드 안 텍스트는 \uFFFC 로 가림."""
    segs, text, depth = [], "", 0
    for r in p.iter(W + "r"):
        anc = r.getparent()
        in_simple = False
        while anc is not None and anc is not p:
            if anc.tag in (W + "fldSimple", W + "del", W + "moveFrom"):
                in_simple = True
                break
            anc = anc.getparent()
        fc = r.find(W + "fldChar")
        if fc is not None:
            t = fc.get(W + "fldCharType")
            depth += 1 if t == "begin" else (-1 if t == "end" else 0)
            continue
        if r.find(W + "instrText") is not None:
            continue
        t = run_text(r)
        if not t:
            continue
        if depth > 0 or in_simple:
            t = "\uFFFC" * len(t)
        segs.append((r, len(text), len(text) + len(t)))
        text += t
    return segs, text


def split_run(r, k):
    """run r 을 텍스트 k 번째 글자에서 둘로 나누고 (왼쪽, 오른쪽) 반환. 단순 텍스트 run 만 지원."""
    ts = r.findall(W + "t")
    ok_tags = (W + "rPr", W + "t", W + "lastRenderedPageBreak")
    if len(ts) != 1 or any(ch.tag not in ok_tags for ch in r):
        return None
    t = ts[0].text or ""
    if k <= 0 or k >= len(t):
        return None
    right = copy.deepcopy(r)
    for e in right.findall(W + "lastRenderedPageBreak"):
        right.remove(e)
    ts[0].text = t[:k]
    ts[0].set(XML_SPACE, "preserve")
    rt = right.find(W + "t")
    rt.text = t[k:]
    rt.set(XML_SPACE, "preserve")
    r.addnext(right)
    return r, right


def make_run(rpr, kind, text=None):
    r = Element("r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if kind == "instr":
        it = SubElement(r, "instrText")
        it.set(XML_SPACE, "preserve")
        it.text = text
    else:
        fc = SubElement(r, "fldChar")
        fc.set(W + "fldCharType", kind)
    return r


def Element(tag):
    from lxml import etree
    return etree.Element(W + tag)


def SubElement(parent, tag):
    from lxml import etree
    return etree.SubElement(parent, W + tag)


def wrap_field(p, start, end, instr):
    """문단 p 의 텍스트 [start, end) 범위를 Word 필드로 감쌈. 성공 시 True."""
    for _ in range(3):
        segs, text = para_map(p)
        first = next((s for s in segs if s[1] <= start < s[2]), None)
        last = next((s for s in segs if s[1] < end <= s[2]), None)
        if not first or not last:
            return False
        if first[1] < start:
            if not split_run(first[0], start - first[1]):
                return False
            continue
        if last[2] > end:
            if not split_run(last[0], end - last[1]):
                return False
            continue
        break
    segs, text = para_map(p)
    span = [s for s in segs if s[1] >= start and s[2] <= end]
    if not span or span[0][1] != start or span[-1][2] != end:
        return False
    r0, r1 = span[0][0], span[-1][0]
    if r0.getparent() is not r1.getparent():
        return False
    rpr = r0.find(W + "rPr")
    if rpr is not None:  # 필드 기호 run 에는 기울임 등을 넣지 않음
        rpr = copy.deepcopy(rpr)
        for tag in ("i", "iCs", "b", "bCs", "u", "vertAlign"):
            for e in rpr.findall(W + tag):
                rpr.remove(e)
    r0.addprevious(make_run(rpr, "begin"))
    r0.addprevious(make_run(rpr, "instr", instr))
    r0.addprevious(make_run(rpr, "separate"))
    r1.addnext(make_run(rpr, "end"))
    return True


def citation_instr(db, items_with_opts, plain):
    cit_items = []
    for itm, opts in items_with_opts:
        ci = {"id": itm["id"], "uris": db.uris(itm), "itemData": db.csl(itm)}
        ci.update(opts)
        cit_items.append(ci)
    data = {"citationID": rand_id(), "properties": {"formattedCitation": plain, "plainCitation": plain, "noteIndex": 0},
            "citationItems": cit_items,
            "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
    return " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(data, ensure_ascii=False) + " "


# ================================================================== 문서 설정(스타일) 저장
def write_zotero_prefs(docx_path):
    """docProps/custom.xml 에 ZOTERO_PREF_n 을 써서 Human Reproduction 스타일로 지정."""
    xml = (f'<data data-version="3" zotero-version="7.0"><session id="{rand_id()}"/>'
           f'<style id="{STYLE_ID}" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="0"/>'
           '<prefs><pref name="fieldType" value="Field"/>'
           '<pref name="automaticJournalAbbreviations" value="true"/></prefs></data>')
    chunks = [xml[i:i + 255] for i in range(0, len(xml), 255)]
    from xml.sax.saxutils import escape
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        custom = zin.read("docProps/custom.xml").decode("utf-8") if "docProps/custom.xml" in names else None
        props = []
        if custom:
            for m in re.finditer(r"<property\b[^>]*name=\"([^\"]+)\"[^>]*>.*?</property>", custom, re.S):
                if not m.group(1).startswith("ZOTERO_PREF"):
                    props.append(m.group(0))
        pid = 2
        out_props = []
        for pr in props:
            out_props.append(re.sub(r'pid="\d+"', f'pid="{pid}"', pr))
            pid += 1
        for n, ch in enumerate(chunks, 1):
            out_props.append(f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" '
                             f'name="ZOTERO_PREF_{n}"><vt:lpwstr>{escape(ch)}</vt:lpwstr></property>')
            pid += 1
        new_custom = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                      + "".join(out_props) + "</Properties>")
        for nm in names:
            data = zin.read(nm)
            if nm == "docProps/custom.xml":
                continue
            if nm == "[Content_Types].xml" and custom is None:
                s = data.decode("utf-8")
                if "/docProps/custom.xml" not in s:
                    s = s.replace("</Types>", '<Override PartName="/docProps/custom.xml" ContentType='
                                  '"application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
                data = s.encode("utf-8")
            if nm == "_rels/.rels" and custom is None:
                s = data.decode("utf-8")
                if "custom-properties" not in s:
                    s = s.replace("</Relationships>", '<Relationship Id="rIdZoteroCustom" Type="http://schemas.'
                                  'openxmlformats.org/officeDocument/2006/relationships/custom-properties" '
                                  'Target="docProps/custom.xml"/></Relationships>')
                data = s.encode("utf-8")
            zout.writestr(zin.getinfo(nm), data)
        zout.writestr("docProps/custom.xml", new_custom)
    os.replace(tmp, docx_path)


# ================================================================== 메인
def main():
    ap = argparse.ArgumentParser(description="텍스트 인용 (Yuk et al., 2022) → Zotero 필드 자동 연결")
    ap.add_argument("docx", nargs="?", help="입력 .docx (생략하면 파일 선택 창)")
    ap.add_argument("--zotero-dir", default=DEFAULT_ZOTERO_DIR, help=f"zotero.sqlite 가 있는 폴더 (기본: {DEFAULT_ZOTERO_DIR})")
    ap.add_argument("--collection", help="검색할 Zotero 폴더 이름 또는 경로 (예: 'Projects/LNG-IUS')")
    ap.add_argument("-o", "--output", help="출력 파일 경로 (기본: <입력>_zotero.docx)")
    ap.add_argument("--keep-references", action="store_true",
                    help="텍스트 참고문헌 목록을 그대로 두기 (기본: Zotero 참고문헌 필드로 교체)")
    ap.add_argument("--auto", action="store_true", help="묻지 않기: 확실한 것만 연결하고 나머지는 텍스트로 둠")
    ap.add_argument("--yes", action="store_true", help="감지된 폴더와 하위 폴더 포함을 묻지 않고 사용")
    ap.add_argument("--fresh", action="store_true", help="이전 실행에서 저장한 선택(_zotero_map.json)을 무시")
    args = ap.parse_args()

    if not args.docx:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            args.docx = filedialog.askopenfilename(title="원고 .docx 선택", filetypes=[("Word", "*.docx")])
            root.destroy()
        except Exception:
            pass
        if not args.docx:
            sys.exit("입력 .docx 파일을 지정하세요.")
    src = os.path.abspath(args.docx)
    base = os.path.splitext(src)[0]
    out = args.output or base + "_zotero.docx"
    map_path = base + "_zotero_map.json"

    db = ZoteroDB(args.zotero_dir)
    try:
        coll_id, scope = pick_scope(db, args)
        memory = {}
        if os.path.exists(map_path) and not args.fresh:
            try:
                saved = json.load(open(map_path, encoding="utf-8"))
                if saved.get("collection") == coll_id:
                    memory = saved.get("map", {})
                    print(f"(이전 선택 {len(memory)}건을 {os.path.basename(map_path)} 에서 불러왔습니다. 무시하려면 --fresh)\n")
            except Exception:
                pass

        doc = Document(src)
        heading, refs, body_paras = find_ref_section(doc)
        if heading is None:
            print("[주의] 'References' 제목을 찾지 못했습니다. 참고문헌 목록 없이 저자+연도로만 매칭합니다.")
        matcher = Matcher(db, scope, args, memory)
        matcher.match_refs(refs)

        # ---------- 본문 인용 처리
        stats = {"fields": 0, "cites": 0, "unmatched": [], "failed_wrap": [], "suspicious": []}
        cited_keys = set()
        for p in body_paras:
            segs, text = para_map(p)
            if not re.search(r"(19|20)\d{2}", text):
                continue
            jobs = []  # (start, end, [(item|None, opts)], display)
            for m in PAREN_RE.finditer(text):
                inner = m.group(1)
                parts = parse_parts(inner)
                if parts is None:
                    if re.search(r"(?:et al\.?|\band\b)[^;]*?(19|20)\d{2}", inner):
                        stats["suspicious"].append(m.group(0))
                    continue
                jobs.append((m.start(), m.end(), parts, m.group(0), False))
            for m in NARR_RE.finditer(text):
                if m.group("a1").split()[0] in SKIP_WORDS:
                    continue
                ps, pe = m.span("paren")
                if any(s <= ps < e for s, e, *_ in jobs):
                    continue
                ys = [y.strip() for y in m.group("years").split(",")]
                parts = [{"prefix": "", "a1": m.group("a1"), "a2": m.group("a2"), "etal": bool(m.group("etal")),
                          "year": y[:4], "sfx": y[4:]} for y in ys]
                jobs.append((ps, pe, parts, m.group(0), True))
            for start, end, parts, display, narrative in sorted(jobs, key=lambda j: -j[0]):
                resolved = []
                for pt in parts:
                    stats["cites"] += 1
                    one = f"{pt['a1']}{' et al.' if pt['etal'] else (' and ' + pt['a2'] if pt['a2'] else '')}, {pt['year']}{pt['sfx']}"
                    itm = matcher.resolve_cite(pt["a1"], pt["a2"], pt["etal"], pt["year"], pt["sfx"], refs, one)
                    if itm is None:
                        stats["unmatched"].append(one)
                    resolved.append((itm, pt))
                good = [(i, pt) for i, pt in resolved if i is not None]
                if not good:
                    continue
                if len(good) < len(resolved):
                    # 일부만 연결되면 전체를 텍스트로 둠 (필드 결과와 원문이 어긋나지 않도록)
                    stats["failed_wrap"].append(display + "  ← 일부 인용 미연결로 건너뜀")
                    continue
                opts_list = []
                for itm, pt in good:
                    o = {}
                    if pt["prefix"]:
                        o["prefix"] = pt["prefix"] + " "
                    if narrative:
                        o["suppress-author"] = True
                    opts_list.append((itm, o))
                plain = text[start:end]
                if wrap_field(p, start, end, citation_instr(db, opts_list, plain)):
                    stats["fields"] += 1
                    cited_keys |= {i["key"] for i, _ in good}
                else:
                    stats["failed_wrap"].append(display + "  ← Word 서식 구조 때문에 필드 삽입 실패")

        # ---------- 참고문헌 목록 → Zotero 참고문헌 필드
        leftover = []
        if refs and not args.keep_references and stats["fields"]:
            uncited = []
            for r in refs:
                if r.item and r.item["key"] not in cited_keys:
                    uncited.append(db.uris(r.item))  # 표, 그림에서만 인용된 문헌도 목록에 남도록
                    cited_keys.add(r.item["key"])
            leftover = [r for r in refs if not r.item]
            first = refs[0].p
            bib_p = copy.deepcopy(first)
            for ch in list(bib_p):  # 원래 참고문헌 문단의 직접 서식은 버림 → Zotero의 Bibliography 스타일 적용
                bib_p.remove(ch)
            instr = (' ADDIN ZOTERO_BIBL ' + json.dumps({"uncited": uncited, "omitted": [], "custom": []})
                     + ' CSL_BIBLIOGRAPHY ')
            r_text = Element("r")
            t = SubElement(r_text, "t")
            t.text = "{Bibliography: Zotero tab -> Refresh}"
            for el in (make_run(None, "begin"), make_run(None, "instr", instr), make_run(None, "separate"),
                       r_text, make_run(None, "end")):
                bib_p.append(el)
            first.addprevious(bib_p)
            for r in refs:
                if r.item:
                    r.p.getparent().remove(r.p)
            # 사이에 남은 빈 문단 정리
            nxt = bib_p.getnext()
            while nxt is not None and nxt.tag == W + "p" and not "".join(x.text or "" for x in nxt.iter(W + "t")).strip() \
                    and nxt.getnext() is not None and leftover and nxt.getnext() is leftover[0].p:
                nxt.getparent().remove(nxt)
                nxt = bib_p.getnext()

        doc.save(out)
        write_zotero_prefs(out)
        json.dump({"collection": coll_id, "map": memory}, open(map_path, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)

        # ---------- 결과 보고
        print("=" * 72)
        print(f"완료: {out}")
        print(f"  Zotero 필드로 바꾼 인용 묶음: {stats['fields']}개 (개별 인용 {stats['cites']}개 검사)")
        print(f"  참고문헌 연결: {sum(1 for r in refs if r.item)}/{len(refs)}")
        if stats["unmatched"]:
            print(f"  연결 안 된 본문 인용 {len(set(stats['unmatched']))}개 (텍스트로 남김):")
            for u in sorted(set(stats["unmatched"])):
                print("    - " + u)
        if stats["failed_wrap"]:
            print("  필드로 바꾸지 못한 곳:")
            for u in stats["failed_wrap"]:
                print("    - " + u)
        if stats["suspicious"]:
            print("  인용처럼 보이지만 형식이 달라 건너뛴 괄호:")
            for u in sorted(set(stats["suspicious"])):
                print("    - " + short(u, 100))
        if leftover:
            print(f"  Zotero 항목이 없는 참고문헌 {len(leftover)}개는 참고문헌 필드 아래 텍스트로 남겨 두었습니다:")
            for r in leftover:
                print(f"    - #{r.idx} {short(r.text, 90)}")
        print("\n다음 단계: Word에서 결과 파일을 열고 Zotero 탭 → Refresh.")
        print("  (스타일은 Human Reproduction 으로 설정되어 있습니다. 문서 설정 창이 뜨면 Human Reproduction 선택)")
        print(f"  선택 기록: {map_path}  (다시 실행하면 같은 선택을 재사용)")
    finally:
        db.close()


if __name__ == "__main__":
    main()
