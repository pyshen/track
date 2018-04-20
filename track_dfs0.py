import math
import datetime
import win32com.client as w32c

class dm:
    def __init__(self,filename,create = 0):
        #constant
        self.TIME_UNDEF       = 'undefined'
        self.TIME_EQ_REL      = 'Equidistant_Relative'
        self.TIME_NONEQ_REL   = 'Non_Equidistant_Relative'
        self.TIME_EQ_CAL      = 'Equidistant_Calendar'
        self.TIME_NONEQ_CAL   = 'Non_Equidistant_Calendar'
        self.TIMEAXISTYPES = (self.TIME_UNDEF,self.TIME_EQ_REL,self.TIME_NONEQ_REL,self.TIME_EQ_CAL,self.TIME_NONEQ_CAL)
        self.ITEMVALUETYPES = ('Instantaneous','Accumulated','Step_Accumulated','Mean_Step_Accumulated','Reverse_Mean_Step_Accumulated')
        # Variables in object (new/empty file parameters)
        self.TSOPROGID = 'COM.TimeSeries_TSObject'
        self.TSO = 0
        self.TSI = 0
        self.datetime = []
        self.timestepsec = -1
        self.filename = filename
        self.create = create
        #print self.create
    def opendfs(self):
        try:
            self.TSO = w32c.Dispatch("TimeSeries.TSObject")
            self.TSI = w32c.Dispatch("TimeSeries.TSItem")
        except ValueError:
            print "Could not initiate TimeSeries handler."
        self.TSO.Connection.FilePath = self.filename
        
        if self.create & self.TSO.Connection.FileExists:
            print "Overide the old file."
        if not self.create:
            if self.TSO.Connection.FileExists == 'False':
                print "FileNotFound"
            elif self.TSO.Connection.IsFileValid == 'False':
                print "FileNotValid"
            else:
                self.TSO.Connection.Open()
                #self.TSO.Connection.GUIOpen()
    def listeumtypes(self):
        self.eumtypes = self.TSI.GetEumTypes()
        #del self.TSI
    def listeumunits(self):
        self.eumunits = self.TSI.GetEumUnits()
    def getdata(self,itemn):
        args = self.TSO(itemn).GetData()
        return args
    def gettime(self):
        args = self.TSO.Time.GetTime()
        return args
    def filetitle(self,filet):
        self.TSO.Connection.FileTitle = filet
    def startdate(self,time):
        self.TSO.Time.StartTime = time
    def timestep(self,time):
        self.TSO.Time.TimeStep.Year = time[0]
        self.TSO.Time.TimeStep.Month = time[1]
        self.TSO.Time.TimeStep.Day = time[2]
        self.TSO.Time.TimeStep.Hour = time[3]
        self.TSO.Time.TimeStep.Minute = time[4]
        self.TSO.Time.TimeStep.Second = time[5]
        self.TSO.Time.TimeStep.Millisecond = time[6]
    def deletevalue(self,dele):
        self.TSO.DeleteValue = dele
    def addtimesteps(self,nstep):
        self.TSO.Time.AddTimeSteps(nstep)
    def setitemeum(self,itemno,eumtype,eumunit):
        item = self.TSO.Item(itemno)
        item.EumType = eumtype
        item.EumUnit = eumunit
    def additems(self,itemname,eumtype,eumunit,datatype):
        item = self.TSO.NewItem()[0]
        #print itemno
        item.Name = itemname
        item.DataType = datatype
        item.AutoConversion = 'True'
        self.setitemeum(self.TSO.Count,eumtype,eumunit)
    def writeitem(self,itemno,v,data):
        length = len(data)
        for i in range(length):
            self.TSO.Item(itemno).SetDataForTimeStepNr(v[i],data[i])
    def writeitems(self,itemno,data):
        self.TSO.Item(itemno).SetData(data)
    def itemdatatype(self,itemno,datatype):
        self.TSO.Item(itemno).DataType = datatype
    def itemname(self,itemno):
        args = self.TSO.Item(itemno).Name
        return args
    def save(self, force = 0):
        self.filename = self.TSO.Connection.FilePath
        self.TSO.Connection.Save()
        #self.TSO.Connection.GUISave()
    def close(self):
        del self.TSO
class point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def distanceTo(self, p2):
        return ((self.x - p2.x) ** 2 + (self.y - p2.y) ** 2) ** 0.5
    def pointAtDist(self, dist, slope, signX, signY, before = False):
        if slope == float('inf'):
            x = self.x
            if before and signY > 0:
                y = self.y - dist
            else:
                y = self.y + dist
            return point(x, y)
        deltaX = (dist ** 2 / (1 + slope ** 2)) ** 0.5
        if (before and signX > 0) or (not before and signX < 0):
            x = self.x - deltaX
            y = self.y - deltaX * slope
        else:
            x = self.x + deltaX
            y = self.y + deltaX * slope
        return point(x, y)

def calSlope(p1, p2):
    if abs(p1.x - p2.x) < 0.001:
        return float('inf')
    else:
        return (p1.y - p2.y) * 1.0 / (p1.x - p2.x)

def writeTrack(txt, time, point, dist, v):
    txt.write('%10.1f %13.6f %13.6f %10.6f % 10.6f\n' % (time, point.x, point.y, dist, v))

def calculateRamp(starttime, timesteps, deltaT, slope, signX, signY, velocity, omega, rampStartPoint, startDistance, txt, sign):
    for timestep in range(int(timesteps)):
        t = timestep * deltaT
        v = sign * velocity / 2. * math.cos(omega * t) + velocity / 2.
        dist = sign * velocity / 2. * (math.sin(omega * t) / omega + sign * t)
        pt = rampStartPoint.pointAtDist(dist, slope, signX, signY)
        writeTrack(txt, t + starttime, pt, dist + startDistance, v)

def calculateMainTrack(warmupT, timesteps, deltaT, slope, signX, signY, velocity, rampDistance, point1, txt):
    for timestep in range(int(timesteps)):
        t = timestep * deltaT
        pt = point1.pointAtDist(t * velocity, slope, signX, signY)
        dist = t * velocity + rampDistance
        writeTrack(txt, t + warmupT, pt, dist, velocity)

def calculateTrack(x1, y1, x2, y2, deltaT, velocity, rampDistance, outfile):

    point1 = point(x1, y1)
    point2 = point(x2, y2)
    slope = calSlope(point1, point2)
    signX = point2.x - point1.x
    signY = point2.y - point1.y

    # mainTrack time step
    trackDist = point1.distanceTo(point2)
    mainTrackT = trackDist / velocity
    timeStepsMain = round(mainTrackT / deltaT)
    # update trackDist and endPoint(full speed) with the rounded timeSteps
    mainTrackT = timeStepsMain * deltaT
    trackDist = timeStepsMain * deltaT * velocity
    endPoint = point1.pointAtDist(trackDist, slope, signX, signY)
   
   # ramp points
    rampStartPoint = point1.pointAtDist(rampDistance, slope, signX, signY, before = True) 
    warmupT = rampDistance / velocity * 2
    timeStepsRamp = round(warmupT / deltaT)
    # update warmupT with rounded time steps
    warmupT = timeStepsRamp * deltaT
    omega = math.pi / warmupT
    rampDistance = warmupT * velocity / 2
    rampStartPoint = point1.pointAtDist(rampDistance, slope, signX, signY, before = True)
    #ramp down parameters
    rampDownT = warmupT + mainTrackT
    rampDownStartDistance = rampDistance + trackDist
    txt = open(outfile,'w')
    txt.write('%10s %13s %13s %10s %10s\n' % ("Time","X-coordinate","Y-coordinate", "Distance","Velocity"))
    # calculate track
    calculateRamp(0, timeStepsRamp, deltaT, slope, signX, signY, velocity, omega, rampStartPoint, 0, txt, -1)
    calculateMainTrack(warmupT, timeStepsMain, deltaT, slope, signX, signY, velocity, rampDistance, point1, txt)
    calculateRamp(rampDownT, timeStepsRamp + 1, deltaT, slope, signX, signY, velocity, omega, endPoint, rampDownStartDistance, txt, 1)
    txt.close()
def writeDfs0(txtfile, deltaT, startTime):
    f = open(txtfile, 'r')
    header = f.readline()
    item1, item2, item3, item4, item5 = header.split()
    timeindex = []
    time = []
    x = []
    y = []
    d = []
    v = []
    n = 0
    for i in f:
        n = n + 1
        m = i.split()
        timeindex.append(n)
        time.append(float(m[0]))
        x.append(float(m[1]))
        y.append(float(m[2]))
        d.append(float(m[3]))
        v.append(float(m[4]))
    dfs = dm(txtfile[:-4] + '.dfs0', 1)
    dfs.opendfs()
    dfs.filetitle('track')
    dfs.startdate(startTime)
    seconds = int(deltaT)
    milliseconds = (deltaT - seconds) * 1000 
    dfs.timestep([0, 0 , 0, 0, 0, seconds, milliseconds])
    dfs.addtimesteps(n)
    dfs.listeumtypes()
    dfs.listeumunits()
    eumunit = 1
    eumtype = dfs.eumtypes.index(u'TimeStep') + 1
    dfs.additems(item1, eumtype, eumunit, 2)
    eumtype = dfs.eumtypes.index(u'Undefined') + 1
    dfs.additems(item2, eumtype, eumunit, 2)
    eumtype = dfs.eumtypes.index(u'Undefined') + 1
    dfs.additems(item3, eumtype, eumunit, 2)
    eumtype = dfs.eumtypes.index(u'Distance') + 1
    dfs.additems(item4, eumtype, eumunit, 2)
    eumtype = dfs.eumtypes.index(u'Velocity Profile') + 1
    dfs.additems(item5, eumtype, eumunit, 2)
    dfs.writeitem(1,timeindex,time)
    dfs.writeitem(2,timeindex,x)                    
    dfs.writeitem(3,timeindex,y)
    dfs.writeitem(4,timeindex,d)
    dfs.writeitem(5,timeindex,v)
    dfs.save()
    dfs.close()
if __name__ == "__main__":
    """ parameters:
    x1, y1 are the coordinates of the start point with full speed
    x2, y2 are the coordinates of the end point with full speed
    deltaT is the time step
    velocity is the full speed of the ship
    rampDistance is the the distance where the speed of the ship increase from 0 to full speed gradually
    outfile is the output file name
    """
    x1 = 600000.0000
    y1 = 3300180.0
    x2 = 602300.8861
    y2 = 3300380.0
    deltaT = 0.5 # sec
    velocity = 4.115
    rampDistance = 500
    startTime = datetime.datetime(2018, 1, 1, 0, 0, 0)
    outfile = 'track.txt'
    calculateTrack(x1, y1, x2, y2, deltaT, velocity, rampDistance, outfile)
    print(outfile + " has been created")
    writeDfs0(outfile, deltaT, startTime)
    print("dfs0 has been created")
