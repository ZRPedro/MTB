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
  
def CheckForZeroDerivatives(errseq):
  '''
  Check for zero derivatives
  '''
  app.PrintInfo('Check for state variable derivatives less than the tolerance for the initial conditions:')
  app.EchoOn()
  inc = app.GetFromStudyCase('ComInc')
  inc.iopt_sim = 'rms'   #Simulation method
  inc.iopt_net = 'sym'   #Network representation: Balanced = 'sym', Unbalanced = 'rst'
  inc.iopt_show = 1      #Verify initial conditions
  inc.iopt_adapt = 1     #Automatic step size adaption
  inc.dtgrd = 0.001      #Electromechanical stepsize
  if inc.iopt_adapt:
    inc.dtgrd_max = 0.01 #Maximum stepsize
  inc.errseq = errseq    #Tolerance value for the initial conditions
  
  incResult = inc.ZeroDerivative()
  
  if incResult == 0:
    app.PrintWarn('At least one state variable has a derivative larger than the tolerance, or the required command options have not been set!\n')
  elif incResult == 1:
    app.PrintInfo('All state variable derivatives are less than the tolerance.\n')
  else:
    app.PrintError('Something went wrong!\n')
  
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

def InitConditions(iopt_net = 'sym', iopt_adapt = 1, dtgrd = 0.001, dtgrd_max = 0.01, errseq = 0.01):
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
        
  app.EchoOff() #To limit the output displayed
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

def StartSimulation():
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
excelOutputPath = script.GetAttribute('ExcelOutputPath')

#######################################################
#Compile automatically all relevant dynamic model types
#######################################################
modelType = script.GetAttribute('ModelTypes')
forceRebuild = script.GetAttribute('ForceRebuild')
outputLevel = script.GetAttribute('DisplayCompilerMessages')

CompileDynamicModelTypes(modelType, forceRebuild, outputLevel)

########################################################################################
#Check for state variable derivatives less than the tolerance for the initial conditions  
########################################################################################
initCondTolerance = script.GetAttribute('MaximumError') #Tolerance value for the initial conditions

CheckForZeroDerivatives(initCondTolerance)

############################################
#Calculate a balanced loadflow, and flat run
############################################
app.PrintInfo('Check for a balanced loadflow, and flat run:\n')

LoadFlow(ldfCalcMethod['balanced'])

#Calculate initial conditions
InitConditions(incNetRepres['balanced'])

#Start RMS Simulation
StartSimulation()

###############################################
#Calculate an unbalanced loadflow, and flat run
###############################################
app.PrintInfo('Check for an unbalanced loadflow, and flat run:\n')

LoadFlow(ldfCalcMethod['unbalanced'])

#Calculate initial conditions
InitConditions(incNetRepres['unbalanced'])

#Start RMS Simulation
StartSimulation()    

#############################
#Check for fixed timestep run
#############################
app.PrintInfo('Check for fixed timestep flat run:\n')

LoadFlow(ldfCalcMethod['balanced'])

#Calculate initial conditions
InitConditions(incNetRepres['balanced'], False, 0.001)

#Start RMS Simulation
StartSimulation()       

################################
#Check for variable timestep run
################################
app.PrintInfo('Check for variable timestep flat run:\n')

LoadFlow(ldfCalcMethod['balanced'])

#Calculate initial conditions
InitConditions(incNetRepres['balanced'], True, 0.001, 0.01)

#Start RMS Simulation
StartSimulation()       
