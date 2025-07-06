import re

text = '''
=CONCAT("DN";TEXTO(C17;"# #/#");"""-";D17;"-";E17;"-";F17;"-";G17)
=CONCAT("DN";TEXTO(C17;"# #/#");"""-";D17;"-";E17;"-";F17;"-";G17)
=CONCAT("DN";TEXTO(C59;"# #/#");"""-";D59;"-";E59;"-";F59;"-";G59)
'''

# Replace any pattern with a letter followed by two digits
ABC = [
    ("C", "A"),
    ("D", "B"),
    ("E", "C"),
    ("F", "D"),
    ("G", "E"),
]

for item in ABC:
    print(fr'\b{item[0]}(\d{2})\b', fr'{item[1]}\1')
    text = re.sub(r'\b'+item[0]+r'(\d{2})\b', item[1]+r'\1', text)

print(text)