import os

# Import the Win32 COM client
try:
	import win32com.client
except ImportError:
	raise ImportError, 'This program requires the pywin32 extensions for Python'

import pywintypes # to handle COM errors.

# Import AutoIT (first try)
autoit = None
try:
	autoit = win32com.client.Dispatch("AutoItX3.Control")
except pywintypes.com_error:
	# If can't instanciate, try to register COM control again:
	os.system("regsvr32 /s AutoItX3.dll")

# Import AutoIT (second try if necessary)
if not autoit:
	try:
		autoit = win32com.client.Dispatch("AutoItX3.Control")
	except pywintypes.com_error:
		raise ImportError, "Could not instanciate AutoIT COM module because",e

if not autoit:
	print "Could not instanciate AutoIT COM module."
	sys.exit(1)
		
#Your AutoIt commands go here
#Eg. Use autoit.Sleep("3000") to pause your code for 3 seconds
autoit.Run("notepad.exe")

