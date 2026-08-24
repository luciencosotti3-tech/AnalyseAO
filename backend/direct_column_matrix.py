# -*- coding: utf-8 -*-
"""Comparaison directe DCE/ACT colonne par colonne, sans mutation Excel."""
from __future__ import annotations
from collections import Counter
import math, re, unicodedata
from typing import Any
VERSION="1.0"
ROLES=("designation","unit","quantity","unit_price","amount","vat")
NUMERIC_ROLES={"quantity","unit_price","amount","vat"}

def norm(value:Any)->str:
    if value is None:return ""
    text=unicodedata.normalize("NFKD",str(value));text="".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+"," ",re.sub(r"[^a-z0-9%]+"," ",text.casefold().replace("²","2"))).strip()

def number(value):
    if value is None or isinstance(value,bool):return None
    if isinstance(value,(int,float)):
        if isinstance(value,float) and math.isnan(value):return None
        return float(value)
    text=str(value).strip().replace("€","").replace("\u00a0","").replace(" ","")
    if not text:return None
    if "," in text and "." in text:text=text.replace(".","").replace(",",".")
    else:text=text.replace(",",".")
    try:return float(text)
    except ValueError:return None

def pure_blocks(ws):
    ribbon=None
    for row in range(1,min(ws.max_row,20)+1):
        if any(norm(c.value)=="estimation" for c in ws[row]):ribbon=row;break
    if ribbon is None:return None,[]
    blocks=[]
    for rg in ws.merged_cells.ranges:
        if rg.min_row==ribbon and rg.max_row==ribbon:
            label=str(ws.cell(ribbon,rg.min_col).value or "").strip()
            if label:blocks.append((label,rg.min_col,rg.max_col))
    return ribbon,sorted(blocks,key=lambda x:x[1])

def role_columns(ws,start,end,ribbon):
    result={role:[] for role in ROLES}
    upper=min(ws.max_row,ribbon+15)
    for col in range(start,end+1):
        parts=[]
        for row in range(ribbon+1,upper+1):
            value=ws.cell(row,col).value
            if value not in (None,""):parts.append(norm(value))
        joined=" ".join(parts);last=next((x for x in reversed(parts) if x),"")
        if "designation" in joined:result["designation"].append(col)
        elif last in {"u","un","unite"}:result["unit"].append(col)
        elif "quantite" in joined or "qte" in joined or "qty" in joined:result["quantity"].append(col)
        elif last in {"p u","pu","prix unitaire"} or "prix unitaire" in joined:result["unit_price"].append(col)
        elif "montant" in joined and "tva" not in joined:result["amount"].append(col)
        elif "tva" in joined:result["vat"].append(col)
    return result

def is_data_row(ws,row,est_roles):
    dcols=est_roles.get("designation",[]);designation=next((ws.cell(row,c).value for c in dcols if ws.cell(row,c).value not in (None,"")),None)
    text=norm(designation)
    if not text or any(x in text for x in ("total","tva","ttc","montant ht")):return False
    for role in ("quantity","unit_price","amount"):
        if any(number(ws.cell(row,c).value) is not None for c in est_roles.get(role,[])):return True
    return False

def equal(a,b,role,tolerance=0.02):
    if role in NUMERIC_ROLES:
        na,nb=number(a),number(b)
        if na is not None and nb is not None:return abs(na-nb)<=tolerance,"EQUAL" if abs(na-nb)<=tolerance else "DIFFERENT"
        if na is None and nb is None:
            return norm(a)==norm(b),"EQUAL" if norm(a)==norm(b) else "TEXT_DIFFERENT"
        return False,"TYPE_DIFFERENT"
    return norm(a)==norm(b),"EQUAL" if norm(a)==norm(b) else "DIFFERENT"

def compare(ws):
    ribbon,blocks=pure_blocks(ws)
    if ribbon is None or len(blocks)<2:raise ValueError("rubans estimation/entreprises introuvables")
    est_label,est_start,est_end=blocks[0];est_roles=role_columns(ws,est_start,est_end,ribbon)
    comparisons=Counter();companies=[];samples=[];paired_columns=0;unmatched_dce=0;unmatched_act=0;rows=0
    for label,start,end in blocks[1:]:
        company=str(label).splitlines()[0].strip();act_roles=role_columns(ws,start,end,ribbon);company_counts=Counter();column_pairs=[]
        for role in ROLES:
            dcols=est_roles[role];acols=act_roles[role];count=min(len(dcols),len(acols));paired_columns+=count;unmatched_dce+=max(0,len(dcols)-count);unmatched_act+=max(0,len(acols)-count)
            column_pairs.extend((role,index+1,dcols[index],acols[index]) for index in range(count))
        for row in range(ribbon+1,ws.max_row+1):
            if not is_data_row(ws,row,est_roles):continue
            rows+=1
            for role,slot,dcol,acol in column_pairs:
                a=ws.cell(row,dcol).value;b=ws.cell(row,acol).value;ok,status=equal(a,b,role);comparisons[status]+=1;company_counts[status]+=1
                if not ok and len(samples)<20:samples.append({"company":company,"row":row,"role":role,"slot":slot,"dce_cell":ws.cell(row,dcol).coordinate,"act_cell":ws.cell(row,acol).coordinate,"dce_value":a,"act_value":b,"status":status})
        companies.append({"company":company,"roles_dce":{k:len(v) for k,v in est_roles.items()},"roles_act":{k:len(v) for k,v in act_roles.items()},"counts":dict(company_counts)})
    return {"version":VERSION,"read_only":True,"method":"ROLE_OCCURRENCE_COLUMN_BY_COLUMN","sheet":ws.title,"company_count":len(blocks)-1,"paired_columns":paired_columns,"unmatched_dce_columns":unmatched_dce,"unmatched_act_columns":unmatched_act,"row_passes":rows,"comparison_counts":dict(comparisons),"companies":companies,"sample_differences":samples}
