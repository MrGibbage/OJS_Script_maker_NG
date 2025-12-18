import ast, sys
p = "c:\\github-projects\\OJS_Script_maker_NG\\2025-all\\build-tournament-folders.py"
src = open(p, "r", encoding="utf-8").read()
tree = ast.parse(src)
defs = {n.name for n in tree.body if isinstance(n, ast.FunctionDef)}
refs = set()
class Finder(ast.NodeVisitor):
    def visit_Name(self, node): refs.add(node.id)
Finder().visit(tree)
unused = sorted([f for f in defs if f not in refs])
print("unused functions:", unused)