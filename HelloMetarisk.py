#
# HelloMetarisk.py -- sample MetaRisk access via IronPython 
#

import metarisk
#from metarisk import *

lossCauseName = 'Florida'
projectName = 'test-project.xmr'
model = metarisk.UserModelRequestProcessor()

newVariation = model.NewProject()
print(f"Adding loss cause: {lossCauseName}")
losscause = model.Add(metarisk.LossCauseDto(lossCauseName), newVariation)
losscause["Severity"]["Attritional Mean"] = 10
model.SaveProject(projectName)
