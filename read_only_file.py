
file_path = 'serial_data.txt'

with open (file_path, 'r') as file:
    for line in file:
        if line.startswith('MAC'):
         print(line)
