import math

def inst_current(Irms, AngBeta):
    #angBeta = AngBeta * 180/math.pi 
    
    #convert rms current to maximum current
    Imax = Irms * math.sqrt(2)
    return str(round(Imax, 2)) + "cos(\u03C9t + " + str(AngBeta) +")"

def inst_voltage(Vrms, AnglDelta):
    #angDelta = AngDelta * 180/math.pi

    Vmax = Vrms * math.sqrt(2)
    return str(round(Vmax, 2)) + "cos(\u03C9t + " + str(AnglDelta) + ")"

def ave_power(Vrms, Irms, AngDelta=0, AngBeta=0):

    if AngBeta != 0 or AngDelta != 0:
        return str(Vrms * Irms) + "cos(" + str(AngDelta - AngBeta) + ")"
    else:
        return Irms * Vrms

def pf_voltage_current(Vrms=0, Irms=0, AngDelta=0, AngBeta=0):
    #convert angle radian to degrees
    angBeta = AngBeta * math.pi/180
    angDelta = AngDelta * math.pi/180

    return round(math.cos(angDelta - angBeta), 3)

def pf_by_z(resistance=0, reactance=0):
    #gets the power factor of a load or impedance given the resistance and the imaginary part reactance

    return round(math.cos(math.atan(reactance/resistance)), 3)

def mag_Z(resistance=0, reactance=0):

    return round(math.sqrt(resistance**2 + reactance**2), 3)

print(pf_voltage_current(230, 15, 23, 67))
print(pf_by_z(9, 6))
print(mag_Z(9, 6))
