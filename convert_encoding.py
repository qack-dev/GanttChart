import codecs

file_path = 'Y:/prog/VBA/Excel_XVBA/GanttChart/vba-files/Class/C_Task.cls'

with codecs.open(file_path, 'r', 'utf-8') as f:
    content = f.read()

with codecs.open(file_path, 'w', 'shift_jis') as f:
    f.write(content)
