#
# HelloMetarisk.py -- sample MetaRisk access via IronPython 
#

import metarisk
#from metarisk import *

lossCauseName = 'Florida'
projectname = 'test-project.xmr'
ecm = metarisk.CapitalModel("./", projectname)
newVariation = ecm.newproject()
print(f"Adding loss cause: {lossCauseName}")
ecm.addlosscause("*", lossCauseName)
ecm.saveproject()