'''
Multiple layers exporter:
    For each layer: 1 textframe for naming, 1 group for graphic content.
'''

print('''
Requirements for exporing multiple layers:
    For each layer: only 1 TextFrame for naming the exported file, only 1 Group for graphic contents.
''')

files_path = input('Path of files to be processed: ')
base_path = input('Base path: ')

import os
files = []
for file in os.listdir(files_path):
    if '.ai' in file:
        files.append(file)
print(f'Deteceted {len(files)} files.')

from handle_file import AIFileHandler
handler = AIFileHandler(files_path, base_path)
for file in files:
    handler.process_file(file)
