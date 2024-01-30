filename = r'D:\IT\Backup sql\backup08.02.2023.sql'
with open(filename, 'r') as file:
    contents = file.read()

with open(filename, 'w') as file:
    file.write('USE `buildte1_cad2`;\n' + contents)
