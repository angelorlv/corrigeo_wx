import language_tool_python

tool = language_tool_python.LanguageTool('fr') 
is_bad_rule = lambda rule: rule.message == 'Faute de frappe possible trouvée.' and len(rule.replacements) and rule.replacements[0][0].isupper()

text = "La LOI_2003-011_FONCTIONNAIRE s'applique aux agent de l'Etat qui occupent des emplois"

matches = tool.check(text)
matches = [rule for rule in matches if not is_bad_rule(rule)]
print(matches)
corrrige = language_tool_python.utils.correct(text, matches)
print(corrrige)