# -*- coding: utf-8 -*-
"""
Created on Fri Feb 21 10:33:29 2025

@author: PRW
"""

import powerfactory

ldfCalcMethod = {
  'balanced'   : 0, #AC Load Flow, balanced, positive sequence
  'unbalanced' : 1, #AC Load Flow, unbalanced, 3-phase (ABC)
  'dc'         : 2  #DC Load Flow linear
  }

incNetRepres = {
  'balanced'   : 'sym', #Balanced, positive sequence
  'unbalanced' : 'rst',  #Unbalanced, 3-phase (ABC)
  }

  
def CompileDynamicModelTypes(modelType, forceRebuild, outputLevel):
  '''
  Try to complile all Dynamic Model Types specified
  '''
  app.PrintInfo('Compile automatically all relevant dynamic model types.')
  app.EchoOn()
  inc = app.GetFromStudyCase('ComInc')
  compileResult = inc.CompileDynamicModelTypes(modelType, forceRebuild, outputLevel)
    
  if compileResult == 0:
    app.PrintInfo('Success!\n')
  elif compileResult == 1:
    app.PrintWarn('Success, but some DSL Model Types will run interpreted.\n')
  elif compileResult == 2:
    app.PrintError('Error!\n')
  else:
    app.PrintError('Something went wrong...\n')

    
def LoadFlow(iopt_net = 0, ):
  '''
  Perform a Load Flow
  '''
  if iopt_net == 0:
    app.PrintInfo('Performing a balanced, positive sequence loadflow...')
  elif iopt_net == 1:
    app.PrintInfo('Performing an unbalanced, 3-phase (ABC) loadflow...')
  else:
    app.PrintError('Unknown value for "iopt_net"!')
    exit(1)
    
  app.EchoOff() #To limit the output displayed
  ldf = app.GetFromStudyCase("ComLdf")
  ldf.iopt_net = iopt_net 
  ldfResults=ldf.Execute()
  if ldfResults == 0:
    app.PrintInfo('Success!\n')
  else:  
    app.PrintError('Non convergence of loadflow analysis.\n')


def InitConditions(iopt_net = 'sym', iopt_adapt = True, dtgrd = 0.001, dtgrd_max = 0.01, errseq = 0.01):
  '''
  Calculated Initial Conditions
  '''
  if iopt_net == 'sym':
    app.PrintInfo('Calculating initial conditions for a balanced, positive sequence network...')
  elif iopt_net == 'rst':
    app.PrintInfo('Calculating initial conditions for an unbalanced, 3-phase (ABC) network...')
  else:
    app.PrintError('Unknown value for "iopt_net"!')
    exit(1)
        
  app.EchoOn() #To limit the output displayed
  inc = app.GetFromStudyCase('ComInc')
  #Fixed
  inc.iopt_sim = 'rms' #Simulation method
  inc.iopt_show = 1    #Verify initial conditions
  #Adjustable
  inc.iopt_net = iopt_net     #Network representation: Balanced = 'sym', Unbalanced = 'rst'
  inc.iopt_adapt = iopt_adapt  #Automatic step size adaption
  inc.dtgrd = dtgrd    #Electromechanical stepsize
  if inc.iopt_adapt:
    inc.dtgrd_max = dtgrd_max #Maximum stepsize
  inc.errseq = errseq #Tolerance value for the initial conditions
  
  incResult = inc.Execute()
  if incResult == 0:
    app.PrintInfo('Success!\n')
  else:  
    app.PrintError("Initial conditions could not be calculated.\n")

  app.PrintInfo('Check for state variable derivatives less than the tolerance for the initial conditions:')
  incResult = inc.ZeroDerivative() 
  if incResult == 0:
    app.PrintWarn('At least one state variable has a derivative larger than the "MaximumError" tolerance!\n')
  elif incResult == 1:
    app.PrintInfo('All state variable derivatives are less than the tolerance.\n')
  else:
    app.PrintError('Something went wrong!\n')

 
def StartRMSSimulation():
  '''
  Performing an RMS Simulation
  '''
  app.PrintInfo('Starting RMS Simulation...')
  app.EchoOff() #To limit the output displayed
  sim = app.GetFromStudyCase("ComSim")
  simResult = sim.Execute()
  if simResult == 0:
    app.PrintInfo('RMS Simulation successfully completed.\n')
  else:  
    app.PrintError("Something went wrong during the RMS simulation!\n")  

#####
#Main
#####
app = powerfactory.GetApplication()
app.ClearOutputWindow()
  
script = app.GetCurrentScript()

initCondTolerance = script.GetAttribute('MaximumError') #Tolerance value for the initial conditions

app.PrintInfo('#############################################################')
app.PrintInfo('Compile all relevant dynamic model types (if ForceRebuild==1)')
app.PrintInfo('#############################################################\n')

modelType = script.GetAttribute('ModelTypes')
forceRebuild = script.GetAttribute('ForceRebuild')
outputLevel = script.GetAttribute('DisplayCompilerMessages')
CompileDynamicModelTypes(modelType, forceRebuild, outputLevel)

app.PrintInfo('#################################################')
app.PrintInfo('Check for balanced loadflow and variable timestep')
app.PrintInfo('#################################################\n')

LoadFlow(ldfCalcMethod['balanced'])
InitConditions(incNetRepres['balanced'], iopt_adapt=True, errseq=initCondTolerance)
StartRMSSimulation()

app.PrintInfo('####################################################')
app.PrintInfo('Check for unbalanced loadflow and variable timestep:')
app.PrintInfo('####################################################\n')

LoadFlow(ldfCalcMethod['unbalanced'])
InitConditions(incNetRepres['unbalanced'], iopt_adapt=True, errseq=initCondTolerance)
StartRMSSimulation()    

app.PrintInfo('##################################################')
app.PrintInfo('Check for balanced loadflow with a fixed timestep:')
app.PrintInfo('##################################################\n')

LoadFlow(ldfCalcMethod['balanced'])
InitConditions(incNetRepres['balanced'], iopt_adapt=False, dtgrd=0.001, errseq=initCondTolerance)
StartRMSSimulation()       

app.PrintInfo('####################################################')
app.PrintInfo('Check for unbalanced loadflow with a fixed timestep:')
app.PrintInfo('####################################################\n')

LoadFlow(ldfCalcMethod['unbalanced'])
InitConditions(incNetRepres['unbalanced'], iopt_adapt=False, dtgrd=0.001, errseq=initCondTolerance)
StartRMSSimulation()       
