import sys
import os
from Manager import Manager
from datetime import date

original_stdout = sys.stdout
original_stderr = sys.stderr

output_file_path = './.log/' + date.today().strftime("%d-%m-%Y") + '.log'
file_exists = os.path.exists(output_file_path)

manager = Manager()

with open(output_file_path, 'a' if file_exists else 'w') as log_file:
    sys.stdout = log_file
    sys.stderr = log_file

    manager.run()

sys.stdout = original_stdout
sys.stderr = original_stderr