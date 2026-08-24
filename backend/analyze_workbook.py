# -*- coding: utf-8 -*-
from argparse import ArgumentParser
from offer_analysis import analyse_workbook

p = ArgumentParser(description="Analyse un classeur restructuré AnalyseAO")
p.add_argument("--workbook", required=True)
p.add_argument("--out", required=True)
a = p.parse_args()
out, report = analyse_workbook(a.workbook, a.out)
print(f"Classeur analysé : {out.resolve()}")
for item in report:
    print(f"{item['sheet']} : {item['status']} - {item['issues']} alerte(s)")
