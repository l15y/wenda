# Pyto.app
# note: first implementation to run on ios 
# interpreter Python 3.8+ in iOS
import sys
import runpy

sys.argv.append('test/fullPresentation.md')
sys.argv.append('test.pptx')

runpy.run_path('./md2pptx', init_globals={'sys': sys})
