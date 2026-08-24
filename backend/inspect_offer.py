# -*- coding: utf-8 -*-
from argparse import ArgumentParser
from excel_engine import analyze_offer, export_analysis_workbook

parser = ArgumentParser(description="Analyse DCE / ACT et génération du classeur Excel comparatif")
parser.add_argument("--dce", required=True)
parser.add_argument("--act", required=True)
parser.add_argument("--entreprise", required=True)
parser.add_argument("--lot", required=True)
parser.add_argument("--output", default="analyse_offres.xlsx")
args = parser.parse_args()

result = analyze_offer(args.dce, args.act, args.entreprise, args.lot)
target = export_analysis_workbook(result, args.output)
print(f"DCE : {result['dce']['article_count']} articles")
print(f"ACT : {result['act']['article_count']} articles")
print(f"Alignements : {result['alignment_count']}")
print(f"Anomalies : {result['issue_counts']}")
print(f"Classeur Excel : {target.resolve()}")
