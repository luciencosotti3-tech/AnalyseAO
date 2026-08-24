# -*- coding: utf-8 -*-
"""Diagnostic contextuel V2, sans mutation du classeur."""
from __future__ import annotations
from collections import Counter
import hashlib, re, unicodedata
from typing import Any
VERSION="2.0"

def norm(v:Any)->str:
    s=unicodedata.normalize("NFKD",str(v or "")); s="".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+"," ",re.sub(r"[^a-z0-9%]+"," ",s.casefold().replace("²","2"))).strip()

def slug(v): return re.sub(r"[^a-z0-9]+",".",norm(v)).strip(".") or "unknown"
def digest(*p): return hashlib.sha1("|".join(str(x or "") for x in p).encode("utf-8","replace")).hexdigest()[:16]

def category(message):
    t=norm(message)
    if t.startswith("calcul incoherent"): return "CALCULATION_ERROR"
    if "prix unitaire manquant" in t or "poste non valorise" in t: return "UNVALUED_OR_TEXTUAL"
    if "quantite estimation" in t and "quantite entreprise" in t: return "QUANTITY_CHANGED"
    if "designation modifiee" in t: return "DESIGNATION_CHANGED"
    if "unite modifiee" in t: return "UNIT_CHANGED"
    if "superieur de plus de 30" in t: return "ABOVE_ESTIMATE_30"
    if "inferieur de plus de 30" in t: return "BELOW_ESTIMATE_30"
    return "COMMENT_TO_REVIEW"

def severity(cat):
    if cat in {"CALCULATION_ERROR","UNVALUED_OR_TEXTUAL"}: return "ERROR"
    if cat in {"ABOVE_ESTIMATE_30","BELOW_ESTIMATE_30"}: return "WARNING"
    return "NOTICE"

def merged_value(ws,row,col):
    value=ws.cell(row,col).value
    if value not in (None,""): return value
    for rg in ws.merged_cells.ranges:
        if rg.min_row<=row<=rg.max_row and rg.min_col<=col<=rg.max_col:
            return ws.cell(rg.min_row,rg.min_col).value
    return value

def header(ws,col,ribbon_row):
    parts=[]
    for row in range(ribbon_row+1,min(ws.max_row,ribbon_row+15)+1):
        value=merged_value(ws,row,col)
        if value not in (None,""): parts.append(str(value))
    return norm(" ".join(parts))

def scope(text):
    m=re.search(r"\bbatiment\s+([a-z0-9]+)\b",norm(text))
    if m:return m.group(1).upper()
    if "degh" in norm(text): return "DEGH"
    if "total" in norm(text) or "cumul" in norm(text): return "TOTAL"
    return "UNKNOWN"

def designation_col(ws,start,end,ribbon_row):
    for row in range(ribbon_row+1,min(ws.max_row,ribbon_row+15)+1):
        for col in range(start,end+1):
            if norm(ws.cell(row,col).value)=="designation": return col
    return min(end,start+1)

def identity(ws,row,start,end,dcol):
    designation=str(ws.cell(row,dcol).value or "").strip()
    reference=""
    for col in range(start,dcol):
        raw=ws.cell(row,col).value
        if raw not in (None,""): reference=str(raw).strip(); break
    return reference,designation,f"{slug(reference)}|{slug(designation)}"

def build_registry(ws,ribbon_row,blocks,historical_issues=0):
    records={}; raw_lines=0; duplicate_lines=0; scopes=Counter(); cats=Counter(); sev=Counter(); companies=Counter()
    for label,start,end in blocks[1:]:
        company=str(label).splitlines()[0].strip(); dcol=designation_col(ws,start,end,ribbon_row); occurrences=Counter()
        headers={col:header(ws,col,ribbon_row) for col in range(start,end+1)}
        for row in range(ribbon_row+1,ws.max_row+1):
            ref,des,base=identity(ws,row,start,end,dcol); occurrences[base]+=1
            row_key=f"{slug(ws.title)}|{base}|occurrence.{occurrences[base]}"
            for col in range(start,end+1):
                comment=ws.cell(row,col).comment
                if not comment or not comment.text: continue
                for raw in comment.text.splitlines():
                    message=raw.strip()
                    if not message: continue
                    raw_lines+=1; cat=category(message); sc=scope(headers[col])
                    issue_id=digest(ws.title,company,row_key,sc,cat,norm(message))
                    if issue_id in records: duplicate_lines+=1; continue
                    record={"issue_id":issue_id,"sheet":ws.title,"company":company,"row_key":row_key,
                            "source_row":row,"source_cell":ws.cell(row,col).coordinate,"category":cat,
                            "severity":severity(cat),"message":message,"building_scope":sc,
                            "reference":ref,"designation":des}
                    records[issue_id]=record; cats[cat]+=1; sev[record["severity"]]+=1; scopes[sc]+=1; companies[company]+=1
    unique=len(records); unknown=scopes.get("UNKNOWN",0)
    return {"version":VERSION,"read_only":True,"historical_issues":int(historical_issues or 0),
            "raw_comment_lines":raw_lines,"unique_records":unique,"duplicates_ignored":duplicate_lines,
            "historical_minus_unique":int(historical_issues or 0)-unique,"unknown_scope_records":unknown,
            "scope_coverage_percent":round((unique-unknown)*100/unique,1) if unique else 100.0,
            "categories":dict(sorted(cats.items())),"severities":dict(sorted(sev.items())),
            "scopes":dict(sorted(scopes.items())),"companies":dict(sorted(companies.items())),
            "sample_records":list(records.values())[:12]}
