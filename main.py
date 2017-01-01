import documents as doc
v = doc.Source()
variables = v.vars_write_to_excel()
t = doc.Target(variables)
t.empty()
replacements = t.read_replacements()
t.compile()








































