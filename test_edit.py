import sys
sys.path.append('.')
from actions.code_helper import _edit_action
print(_edit_action('test.py', 'change hello to bye', None))
