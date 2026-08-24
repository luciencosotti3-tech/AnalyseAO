# -*- coding: utf-8 -*-
from argparse import ArgumentParser
from side_by_side import build_side_by_side

p = ArgumentParser(description="Met les tableaux DCE et ACT côte à côte en conservant leurs structures")
p.add_argument("--dce", required=True)
p.add_argument("--act", required=True)
p.add_argument("--entreprise", required=True)
p.add_argument("--lot", required=True)
p.add_argument("--output", default="comparaison_cote_a_cote.xlsx")
a = p.parse_args()
target = build_side_by_side(a.dce, a.act, a.entreprise, a.lot, a.output)
print(f"Classeur côte à côte : {target.resolve()}")
