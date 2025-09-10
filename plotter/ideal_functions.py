import sys
import numpy as np
import pandas as pd
from scipy.signal import bilinear, lfilter, lfiltic
from Result import ResultType


def genIdealResults(result, resultData, settingDict, caseDf, pscadInitTime):
    '''
    Generates ideal results for different test cases based on PSCAD simulation data and settings.

    Parameters:
        result: Result object containing result type information.
        resultData: pandas.DataFrame containing simulation data.
        settingDict: dict containing test settings and parameters.
        caseDf: pandas.DataFrame containing case configuration and events.
        pscadInitTime: float, initial time offset from PSCAD simulation.

    Returns:
        dict with keys:
            'figs': list of figure names to plot,
            'signals': list of signal names in the DataFrame,
            'data': pandas.DataFrame with calculated ideal results.
    '''
    
    # Use PSCAD result for calculating the ideal response
    if result.typ in (ResultType.EMT_INF, ResultType.EMT_PSOUT, ResultType.EMT_CSV, ResultType.EMT_ZIP):
        idealData = resultData.rename(columns=dict(zip(resultData.columns, [val.split('\\')[-1] for val in resultData.columns])), inplace=False) # Don't set inplace=True, it will also change the original DataFrame
        idealData['time'] = idealData.time - pscadInitTime
        
        # Generic LPF settings
        Ts = idealData['time'].iloc[1]-idealData['time'].iloc[0]  # Sampling time
        trise = 0.5                                               # Rise time [s]
        fc = 0.35/trise                                           # Cut-off frequency [Hz]
        
        tThresh = -1                        # Time threshold [s]            
        mask = idealData['time'] >= tThresh # Only apply from t >= tThresh so as not to use initial transients in PSCAD simulation
        
        idealFigs = ['']
        idealSignals = ['']
        
        # Active Power Ramping cases
        if 'P_step' in caseDf['Case']['Name'].squeeze():
            P0 = caseDf['Initial Settings']['P0'].squeeze()                     # Initial active power setpoint, P0
            assert caseDf['Event 1']['type'].squeeze() == 'Pref'
            Tstep = caseDf['Event 1']['time'].squeeze()
            Pstep = caseDf['Event 1']['X1'].squeeze()
            assert caseDf['Event 1']['X2'].squeeze() == 0.0
            
            idealData['P_pu_PoC_Ramp'] = pd.Series([idealPramp(Pref=P0, Tstep=Tstep, Pstep=Pstep, t=t) for t in idealData.time])
            
            idealFigs.append('Ppoc')
            idealSignals.append('P_pu_PoC_Ramp')

        # LFSM, FSM & RoCoF cases    
        if  'FSM' in caseDf['Case']['Name'].squeeze() or 'RoCoF'in caseDf['Case']['Name'].squeeze() or 'ROCOF'in caseDf['Case']['Name'].squeeze():
            P0 = caseDf['Initial Settings']['P0'].squeeze()                     # Initial active power setpoint, P0
            DK = 1 if settingDict['Area']=='DK1' else 2                         # DK area, either 1 or 2
            DSO = True if settingDict['Un']<110 else False                      # DSO, either Energinet (TSO))
            s_fsm = settingDict['FSM droop']                                    # FSM droop in [%]
            db = settingDict['FSM deadband']                                    # FSM deadband in [Hz]
            FSM = caseDf['Initial Settings']['Pmode'].squeeze() == 'LFSM+FSM'   # FSM mode enabled
            fn = 50                                                             # Nominal frequency [Hz]
            Td = 0.2                                                            # Delay time [s]

            idealData['f_hz_Td'] = delay(idealData['pll_f_hz'], Td, Ts)    # Delayed the 'pll_f_hz' signal                       
            idealData.loc[idealData['time'] < tThresh, 'f_hz_Td'] = fn     # Set values for t < tThresh to fn to eliminate the initialisation transients
            idealData['f_hz_Td_Lpf'] = lpf(idealData['f_hz_Td'], fc, 1/Ts) # Pass the delayed signal through an LPF

            if not 'pstep' in caseDf['Case']['Name'].squeeze(): # idealLFSM does not support active power ramping
                idealData['P_pu_LFSM_FFR'] = pd.Series([idealLFSM(Pref=P0, f=f, DK=DK, FSM=FSM, s_fsm=s_fsm, db=db) for f in idealData['f_hz_Td_Lpf']]) # TODO: Change P0 to mtb_s_pref_pu
                idealFigs.append('Ppoc')
                idealSignals.append('P_pu_LFSM_FFR')                                   
            
            idealData['P_pu_LFSM_Ramp'] = P0 # Create the active power LFSM-Ramp signal to 'populate' below
            idealData['P_pu_LFSM_Ramp'] = idealLFSMRamp(P0=idealData['mtb_s_pref_pu'], Ts=Ts, f=idealData['pll_f_hz'], fTdLpf=idealData['f_hz_Td_Lpf'], P=idealData['P_pu_LFSM_Ramp'], DK=DK, FSM=FSM, s_fsm=s_fsm, db=db)  # TODO: Change P0 to mtb_s_pref_pu
            idealFigs.append('Ppoc')
            idealSignals.append('P_pu_LFSM_Ramp')                                            
                        
        # Q control cases
        if caseDf['Initial Settings']['Qmode'].squeeze() == 'Q':            
            Qref0 = caseDf['Initial Settings']['Qref0'].squeeze()               # Initial reactive power setpoint, when Qmode == 'Q'
            
            idealData['Q_pu_PoC'] = lpf(idealData['mtb_s_qref'], fc, 1/Ts)      # Ideal response == Qref passed through a LPF      
            idealData.loc[idealData['time'] < tThresh, 'Q_pu_PoC_Lpf'] = Qref0  # Set values for t < tThresh
            idealFigs.append('Qpoc')
            idealSignals.append('Q_pu_PoC')  

        # Q(U) control cases
        if caseDf['Initial Settings']['Qmode'].squeeze() == 'Q(U)':
            DK = 1 if settingDict['Area']=='DK1' else 2                         # DK area, either 1 or 2
            DSO = True if settingDict['Un']<110 else False                      # DSO, either Energinet (TSO))
            s_V_droop = settingDict['V droop']                                  # Vdroop value
            #Uref0 = caseDf['Initial Settings']['Qref0'].squeeze()               # Note: If Qmode == 'Q(U)', then Qref0 = Uref0
            Qref0 = 0.0

            idealData['Q_pu_QU'] = Qref0      # Create new signal to populate
            for i, row in idealData.iterrows():
                if row['time'] < tThresh:
                    continue
                QpuQU = idealQU(Uref=row['mtb_s_qref'], Upos=row['fft_pos_Vmag_pu'], s=s_V_droop, Qref=Qref0, DK=DK , DSO=DSO) # Note: If Qmode == 'Q(U)', then 'mtb_s_qref' = Uref
                idealData.loc[i, 'Q_pu_QU'] = QpuQU 

            idealFigs.append('Qpoc')
            idealSignals.append('Q_pu_QU') 
            
            idealData['Q_pu_QU_Lpf'] = lpf(idealData['Q_pu_QU'], fc, 1/Ts)            
            idealFigs.append('Qpoc')
            idealSignals.append('Q_pu_QU_Lpf')  
            
        #PF control mode
        if caseDf['Initial Settings']['Qmode'].squeeze() == 'PF':
            P0 = caseDf['Initial Settings']['P0'].squeeze()                     # Initial active power setpoint, P0
            PFref0 = caseDf['Initial Settings']['Qref0'].squeeze()              # Initial PF setpoint, Qref0, when Qmode == 'PF'
            
            if caseDf['Event 1']['type'].squeeze() == 'Pref':   # Pref changes -> slow ramping of Ppoc, thus use Ppoc and not Pref
                idealData['Q_pu_Qpf'] = idealQpf(Ppoc=idealData['P_pu_PoC'], PFref=PFref0)
            elif caseDf['Event 1']['type'].squeeze() == 'Qref': # PFref changes & Pref constant
                idealData['Q_pu_Qpf'] = idealQpf(Ppoc=idealData['mtb_s_pref_pu'], PFref=idealData['mtb_s_qref']) # Note that .mtb_s_qref = PF if Qmode == 'PF'
            else: # Use Initial settings
                idealData['Q_pu_Qpf'] = idealQpf(Ppoc=P0, PFref=PFref0)
                
            idealData['Q_pu_Qpf_Lpf'] = lpf(idealData['Q_pu_Qpf'], fc, 1/Ts)            
            idealData.loc[idealData['time'] < tThresh, 'Q_pu_Qpf_Lpf'] = P0*np.tan(np.arccos(PFref0))      # Set values for t < tThresh
            idealFigs.append('Qpoc')
            idealSignals.append('Q_pu_Qpf_Lpf')  
                        
        # Fast Fault Current contribution cases
        if 'FRT' in caseDf['Case']['Name'].squeeze() or 'Fault' in caseDf['Case']['Name'].squeeze() or 'support' in caseDf['Case']['Name'].squeeze():
            DK = 1 if settingDict['Area']=='DK1' else 2                         # DK area, either 1 or 2
            DSO = True if settingDict['Un']<110 else False                      # DSO, either Energinet (TSO))
            if DK == 1:
                vposFrtLimit = 0.85                                             # FRT Voltage limit for DK1
            elif DK == 2 or DSO:
                vposFrtLimit = 0.9                                              # FRT Voltage limit for DK2 or DSO cases
                
            Iq0 = 0
            idealData['Iq_pu_FFC'] = 0.0    # Create new signal to populate
            for i, row in idealData.iterrows():
                if row['time'] < tThresh:
                    continue
                if row['fft_pos_Vmag_pu'] >= vposFrtLimit:
                    Iq0 =  row['fft_pos_Iq_pu']
                    #Iq0 = 0 # TODO: Fix that the Iq0 value is used before FRT is detected
                IqFFC = idealFFC(Upos=row['fft_pos_Vmag_pu'], Iq0 = Iq0, DK=DK, DSO=DSO)
                    
                idealData.loc[i, 'Iq_pu_FFC'] = IqFFC                                     
            
            idealFigs.append('Ireactive')
            idealSignals.append('Iq_pu_FFC')  

        returnDict = {'figs': idealFigs, 'signals': idealSignals, 'data': idealData}
            
    else:
        returnDict = {'figs': [''], 'signals': [], 'data': pd.DataFrame()}

    return returnDict


def lpf(x, fc, fs):
    '''
    Simple first order low pass filter with cut off frequency fc and sampling frequency fs

    Parameters:
        x: input signal 
        fc: cut off frequency in Hz
        fs: sampling frequency in Hz
    
    Returns:
        filtered signal 
    '''
    wc = 2*np.pi*fc
    b, a = bilinear([0, wc], [1, wc], fs)
    zi = lfiltic(b, a, y=x, x=x)    # Calculate flfilter's initial condition
    y, zf = lfilter(b, a, x, zi=zi) # zf holds the final filter delay values and is required with 'zi', but not used

    return y


def delay(x, Td, Ts):
    '''
    Simple signal delay of Td seconds with sampling time Ts

    Parameters:
        x: input signal
        Td: delay time in seconds
        Ts: sampling time in seconds

    Returns:
        delayed signal
    '''
    delay_samples = int(round(Td/Ts))
    b = np.zeros(delay_samples + 1)
    b[delay_samples] = 1    # For a pure delay of N samples, b = [0, 0, ..., 0, 1] (where 1 is at index N)
    a = np.array([1.0])     # For an FIR filter (pure delay), a = [1]

    if delay_samples == 0:
        return x  # No delay needed, return original signal
    
    # Apply the filter to the input signal
    return lfilter(b, a, x)


def idealPramp(Pref, Tstep, Pstep, t):
    '''
    This function calculates the ideal or maximum rates of change of 
    active power output (Pramp) in both an up and down direction of for 
    a change in the active power reference for power-generating modules
    based on RfG (EU) 2016/631, 15.6 (e) NC 2025 (Version 4)

    Parameters:
        Pref in [pu] -- for Power Park Modules, Pref is Active Power reference *before ramping* 
        Tstep in [s] -- time step for the ramping
        Pstep in [pu] -- the new reference value of the Active Power *after ramping*   
        t in [s] -- time at which the new value of P is calculated

    Returns:
        Pramp in [pu] -- the new value of P after the ramping
    '''
    if Pstep > Pref:
        m =  0.00333333  # pu/s - equivalent to 0.2 pu/min
        Pramp = Pref if t <= Tstep else m*(t-Tstep) + Pref
        Pramp = Pstep if Pramp >= Pstep else Pramp # Ensure Pramp does not exceed the new reference value (Pstep)
    else:
        m = -0.00333333  # pu/s - equivalent to -0.2 pu/min
        Pramp = Pref if t <= Tstep else m*(t-Tstep) + Pref
        Pramp = Pstep if Pramp <= Pstep else Pramp # Ensure Pramp does not go below the new reference value (Pstep)
        
    return Pramp


def idealLFSMRamp(P0, Ts, f, fTdLpf, P, DK, FSM, s_fsm, db):
    '''
    This function ensures that the active power output (P) does not
    change too fast when the frequency (f) is close to the nominal frequency (fn)
    based on RfG (EU) 2016/631, 13.2 (e) NC 2025 (Version 4)
    
    Parameters:
        P0 in [pu] -- the Active Power reference setpoint 
        Ts in [s] -- sampling time   
        f in [Hz] -- the actual frequency
        fTdLpf in [Hz] -- the delayed and low-pass filtered frequency
        P in [pu] -- the Active Power output to be updated by the function
        DK [1,2] -- the Danish area, either 1 or 2
        FSM  [True, False] -- if True, the function will use the FSM droop and deadband
        s_fsm in [%] -- the FSM droop in [%]
        db in [Hz] -- the FSM deadband in [Hz]
        
    Returns:
        P in [pu] -- the new value of P after ensuring the ramping rate is not too high
    '''
    
    fThresh = 0.04      # Frequency threshold
    PThresh = 0.0001    # Active power threshold
    m = 0.00333333      # pu/s - equivalent to 0.2 pu/min ramping rate
    fn = 50.0           # Nominal frequency in Hz
    
    Pref = P0.iloc[0]           
    for k in range(1, len(P)):
        # If the active power is ramping to fast with the frequency close to the nominal frequency
        if np.abs(f.iloc[k] - fn) < fThresh:    # Ramping active
            if np.abs(P0.iloc[k] - P.iloc[k-1]) > PThresh:
                if P0.iloc[k] - P.iloc[k-1] > 0:  # If P needs to increasing
                    P.iloc[k] = m*Ts + P.iloc[k-1]
                    if P.iloc[k] > P0.iloc[k]:  # Ensure P does not exceed P0
                        P.iloc[k] = P0.iloc[k]
                else: # If P needs to decreasing
                    P.iloc[k] = -m*Ts + P.iloc[k-1]
                    if P.iloc[k] < P0.iloc[k]:  # Ensure P does not go below P0
                        P.iloc[k] = P0.iloc[k]
            else:   # Maintain current value of P
                P.iloc[k] = P.iloc[k-1]
            Pref =  P.iloc[k]
        else:   # LFSM active
            P.iloc[k] = idealLFSM(Pref=Pref, f=fTdLpf.iloc[k-1], DK=DK, FSM=FSM, s_fsm=s_fsm, db=db)
    
    return P


def idealLFSM(Pref, f, DK=1, FSM=False, s_fsm=10, db=0):
    '''
    This function calculates the new value of P given the
    frequency f, according to RfG (EU) 2016/631, 13.2 (a-d) and NC 2025
    (Version 4) for either DK1 or DK2

    Parameters:
        Pref in [pu] -- for Power Park Modules, Pref is the actual Active Power output     
        f in [Hz] -- the actual frequency at the point of connection
        DK in [1,2] -- the Danish area, either 1 or 2
        FSM in [True, False] -- if True, the function will use the FSM droop and deadband
        s_fsm in [%] -- the FSM droop in [%]
        db in [Hz] -- the FSM deadband in [Hz]

        If f > fn (50Hz),
            the output will be the LFSM-O value
        else if f < fn (50Hz),
            it will be the LFSM-U value
        If FSM is True,
            the FSM droop, "s_fsm" has to be specified 
            as well as the deadband, "db" if any

    Returns:
        Pnew in [pu] -- the new value of P after the frequency change
    '''
    Pn = 1.0 #pu
    fn = 50  #Hz
    if DK == 1:
        f1 = 50.2 if f > fn else 49.8
        s = 5
    elif DK == 2:
        f1 = 50.5 if f > fn else 49.5
        s = 4
    else:
        print('"DK" can either be "1" or "2"!')
        return 0
    
    if FSM:
        Pref = idealFSM(Pref, f, DK, s_fsm, db)
        
    if f > fn and f > f1 or f < fn and f < f1:
        Pnew = Pref-100/s*(f-f1)/fn*Pn
    else:
        Pnew = Pref
        
    if Pnew >= 1.0:
        Pnew = 1.0  # Limit Active Power to 1.0 pu
    elif Pnew <= 0.0:
        Pnew = 0.0  # Limit Active Power to 0.0 pu

    return Pnew


def idealFSM(Pref, f, DK=1, s=10, db=0):
    '''
    With Pref in pu, the function calculates the new value of P given the
    frequency f, for the FSM droop, s, according to RfG (EU) 2016/631, 15.2 (d)
    and NC 2025 (Version 4)

    Parameters:
        Pref in [pu] -- for Power Park Modules, Pref is the actual Active Power output and not the nominal power
        f in [Hz] -- the actual frequency at the point of connection
        DK in [1,2] -- the Danish area, either 1 or 2
        s in [%] -- the FSM droop in [%]
        db in [Hz] -- the FSM deadband in [Hz]

    Returns:
        Pnew in [pu] -- the new value of P after the frequency change 
    '''
    Pn = 1.0 #pu
    fn = 50  #Hz
    if DK == 1:
        fRU = 49.8
        fRO = 50.2
    elif DK == 2:
        fRU = 49.5
        fRO = 50.5
    else:
        print('"DK" can either be "1" or "2"!')
        return 0

    f = fRU if f < fRU else f
    f = fRO if f > fRO else f
    
    if f<fn:
        Pnew = Pref-100/s*(f-fn+db)/fn*Pn if f<fn-db else Pref
    else:
        Pnew = Pref-100/s*(f-fn-db)/fn*Pn if f>fn+db else Pref        
    
    # Clamp Pnew to to be not exceed Pref by +/- 10%
    if Pnew > Pref+0.1*Pn: Pnew = Pref+0.1*Pn
    if Pnew < Pref-0.1*Pn: Pnew = Pref-0.1*Pn
    
    return Pnew


def idealQU(Uref, Upos, s, Qref=0, DK=1 , DSO=False):
    '''
    This function calculates the value of the reactive power required, Qpoc_QU, 
    for voltage control mode based on RfG (EU) 2016/631, 21.3 (d) NC 2025
    (Version 4)

    Parameters:
        Uref in [pu] -- the voltage reference at the point of connection
        Upos in [pu] -- the magnitude of the positive sequence voltage at the 
                        point of connection
                        TODO: Legacy suppor for Terminal (LV/MV ?)
        s in [%]     -- the voltage droop of the Q(U) control   
        Qref in [pu] -- the Reactive Power reference (which by default is zero)

    Returns:
        Qpoc_QU in [pu] -- the new value of Q at the point of connection
    '''
    Qnom = 0.33     # TODO: Qnom should be the actual Qnom of the plant

    if DK == 1 and not DSO:
        vposFrtLimit = 0.85
    elif DK == 2 or DSO:
        vposFrtLimit = 0.9
    else:
        print(f"DK = {DK} is not a valid option!")
        sys.exit(1)

    if Upos < vposFrtLimit: # No FRT
        Qpoc_QU = Qnom     # Qpoc_QU reference is is clamped at Qnom 
    else:
        dU = Uref-Upos
        dQ = 100*dU/Uref*Qnom/s
        if Qref + dQ > Qnom:
            Qpoc_QU = Qnom
        elif Qref + dQ < -Qnom:
           Qpoc_QU = -Qnom 
        else:
            Qpoc_QU = Qref + dQ
        
    return Qpoc_QU


def idealQpf(Ppoc, PFref): 
    '''
    This function calculates the ideal reactive power output (Qpoc) based on the active power output (Ppoc)
    and the power factor reference (PFref) according to RfG (EU) 2016/631, 21.3 (d) NC 2025 (Version 4)

    Parameters:
        Ppoc in [pu] -- the actual Active Power output at the point of connection
        PFref in [pu] -- the power factor reference

    Returns:
        Qpoc in [pu] -- the new value of Q at the point of connection
    '''
    theta = np.arccos(PFref)
    Qpoc = np.clip(Ppoc*np.tan(theta), -0.33, 0.33)  # Ensure Qpoc is within [-0.33, 0.33] range
                    
    return Qpoc
    

def idealFFC(Upos, Iq0, DK, DSO):
    '''
    This function calculates the fast fault current (FFC) contribution,
    Id (experimental) and Iq, based on the positive sequence voltage, Upos
    based on RfG (EU) 2016/631, 20.2 (b) NC 2025 (Version 4) for DK1 and DK2
    
    Parameters:    
        Upos in [pu] -- the positive sequence voltage magnitude at the point of connection
        DK in [1,2] -- the Danish area, either 1 or 2
        Iq0 in [pu] -- the Iq value just before FRT is entered 
        DSO in [True, False] -- if True, the function will use the DSO voltage limit of 0.9 pu, otherwise it will use the TSO voltage limit of 0.85 pu for DK1
    Returns:
        tuple of (Id, Iq)
        Id in [pu] -- the "remaining" positive sequence, direct current at the point of connection
        IqFFC in [pu] -- the required positive sequence, quadrature current at the point of connection

    '''
    if DK == 1 and not DSO:
        vposFrtLimit = 0.85
    elif DK == 2 or DSO:
        vposFrtLimit = 0.9
    else:
        print(f"DK = {DK} is not a valid option!")
        sys.exit(1)
       
    if Upos >= vposFrtLimit: # No FRT
        IqFFC = Iq0     # Iq is unchanged
    else:
        if Upos < vposFrtLimit and Upos > 0.5:
            if DK==2 or DSO:
                IqFFC = -1/0.4*Upos+2.25 + Iq0
            else: # DK==1
                IqFFC = -1/0.35*Upos+2.42857 + Iq0
            
        else: # Upos <= 0.5
            IqFFC = 1.0 + Iq0
        
    return IqFFC


