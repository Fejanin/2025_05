from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule


workbook = load_workbook(filename="hello_world.xlsx")
sheet = workbook.active


red_background = PatternFill(patternType='solid', fgColor="00FF0000", bgColor="00FF0000")
diff_style = DifferentialStyle(fill=red_background)
rule = Rule(type="expression", dxf=diff_style)
rule.formula = ['A1&B1="helloworld!"']
sheet.conditional_formatting.add("H1:H10", rule)


workbook.save(filename="hello_world_append.xlsx")
