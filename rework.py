import servicemanager
import sys
import win32event
import win32service
import win32serviceutil
import win32timezone
from time import sleep
try:
    import http.client as client
except:
    import httplib as client 

from time import sleep
from socket import *
import select
import errno
import time
from struct import pack, unpack
from datetime import datetime, date,timedelta
import json
#zkconst 


USHRT_MAX = 65535


CMD_CONNECT = 1000
CMD_EXIT = 1001
CMD_ENABLEDEVICE = 1002
CMD_DISABLEDEVICE = 1003

CMD_ACK_OK = 2000
CMD_ACK_ERROR = 2001
CMD_ACK_DATA = 2002

CMD_PREPARE_DATA = 1500
CMD_DATA = 1501

CMD_USERTEMP_RRQ = 9
CMD_ATTLOG_RRQ = 13
CMD_CLEAR_DATA = 14
CMD_CLEAR_ATTLOG = 15

CMD_WRITE_LCD = 66

CMD_GET_TIME  = 201
CMD_SET_TIME  = 202

CMD_VERSION = 1100
CMD_DEVICE = 11

CMD_CLEAR_ADMIN = 20
CMD_SET_USER = 8

LEVEL_USER = 0
LEVEL_ADMIN = 14

def encode_time(t):
    """Encode a timestamp send at the timeclock

    copied from zkemsdk.c - EncodeTime"""
    d = ( (t.year % 100) * 12 * 31 + ((t.month - 1) * 31) + t.day - 1) *\
         (24 * 60 * 60) + (t.hour * 60 + t.minute) * 60 + t.second

    return d


def decode_time(t):
    """Decode a timestamp retrieved from the timeclock

    copied from zkemsdk.c - DecodeTime"""
    second = t % 60
    t = t / 60

    minute = t % 60
    t = t / 60

    hour = t % 24
    t = t / 24

    day = t % 31+1
    t = t / 31

    month = t % 12+1
    t = t / 12

    year = t + 2000

    d = datetime(year, month, day, hour, minute, second)

    return d
    


#zkconst end
#zkconnect 

def zkconnect(self):
    """Start a connection with the time clock"""
    command = CMD_CONNECT
    command_string = ''
    chksum = 0
    session_id = 0
    reply_id = -1 + USHRT_MAX

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    
    self.zkclient.sendto(buf, self.address)
    
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        
        return self.checkValid( self.data_recv )
    except:
        return False
    

def zkdisconnect(self):
    """Disconnect from the clock"""
    command = CMD_EXIT
    command_string = ''
    chksum = 0
    session_id = self.session_id
    
    reply_id = unpack('HHHH', self.data_recv[:8])[3]
    
    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)

    self.zkclient.sendto(buf, self.address)
    
    self.data_recv, addr = self.zkclient.recvfrom(1024)
    return self.checkValid( self.data_recv )
    


#zkconnect end
#zkversion 

def zkversion(self):
    """Start a connection with the time clock"""
    command = CMD_VERSION
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkversion end
#zkkos 

def zkos(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~OS'
    chksum = 0
    session_id = self.session_id
    
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkkos end
#zkextendfmt

def zkextendfmt(self):
    try:
        test = self.exttrynumber
    except:
        self.exttrynumber = 1
        
    data_seq=[ self.data_recv.encode("hex")[4:6], self.data_recv.encode("hex")[6:8] ]
    print(data_seq)
    if self.exttrynumber == 1:
        plus1 = 0
        plus2 = 0
    else:
        plus1 = -1
        plus2 = +1
        
    
    desc = ": +"+hex( int('99', 16)+plus1 ).lstrip('0x')+", +"+hex(int('b1', 16)+plus2).lstrip("0x")
    self.data_seq1 = hex( int( data_seq[0], 16 ) + int( '99', 16 ) + plus1 ).lstrip("0x")
    self.data_seq2 = hex( int( data_seq[1], 16 ) + int( 'b1', 16 ) + plus2 ).lstrip("0x")
    
    if len(self.data_seq1) >= 3:
        #self.data_seq2 = hex( int( self.data_seq2, 16 ) + int( self.data_seq1[:1], 16) ).lstrip("0x")
        self.data_seq1 = self.data_seq1[-2:]
        
    if len(self.data_seq2) >= 3:
        #self.data_seq1 = hex( int( self.data_seq1, 16 ) + int( self.data_seq2[:1], 16) ).lstrip("0x")
        self.data_seq2 = self.data_seq2[-2:]
        

    if len(self.data_seq1) <= 1:
        self.data_seq1 = "0"+self.data_seq1
        
    if len(self.data_seq2) <= 1:
        self.data_seq2 = "0"+self.data_seq2
    
    
    counter = hex( self.counter ).lstrip("0x")
    if len(counter):
        counter = "0" + counter
    print(self.data_seq1+" "+self.data_seq2+desc)
    data = "0b00"+self.data_seq1+self.data_seq2+self.id_com+counter+"007e457874656e64466d7400"
    self.zkclient.sendto(data.decode("hex"), self.address)
    print(data)
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
    except:
        if self.exttrynumber == 1:
            self.exttrynumber = 2
            tmp = zkextendfmt(self)
        if len(tmp) < 1:
            self.exttrynumber = 1
    
    self.id_com = self.data_recv.encode("hex")[8:12]
    self.counter = self.counter+1
    print(self.data_recv.encode("hex"))
    return self.data_recv[8:]


#zkextendfmt end
#zkextendoplog

def zkextendoplog(self, index=0):
    try:
        test = self.extlogtrynumber
    except:
        self.extlogtrynumber = 1
        
    data_seq=[ self.data_recv.encode("hex")[4:6], self.data_recv.encode("hex")[6:8] ]
    print(data_seq)
    
    if index==0:
        self.data_seq1 = hex( int( data_seq[0], 16 ) + int( '104', 16 ) ).lstrip("0x")
        self.data_seq2 = hex( int( data_seq[1], 16 ) + int( '19', 16 ) ).lstrip("0x")
        desc = ": +104, +19" 
        header="0b00"
    elif index==1:
        self.data_seq1 = hex( abs( int( data_seq[0], 16 ) - int( '2c', 16 ) ) ).lstrip("0x")
        self.data_seq2 = hex( abs( int( data_seq[1], 16 ) - int( '2', 16 ) ) ).lstrip("0x")
        desc = ": -2c, -2" 
        header="d107"
    elif index>=2:
        self.data_seq1 = hex( abs( int( data_seq[0], 16 ) - int( '2c', 16 ) ) ).lstrip("0x")
        self.data_seq2 = hex( abs( int( data_seq[1], 16 ) - int( '2', 16 ) ) ).lstrip("0x")
        desc = ": -2c, -2" 
        header="ffff"
    
    
    print(self.data_seq1+"  "+self.data_seq2)
    if len(self.data_seq1) >= 3:
        self.data_seq2 = hex( int( self.data_seq2, 16 ) + int( self.data_seq1[:1], 16) ).lstrip("0x")
        self.data_seq1 = self.data_seq1[-2:]
        
    if len(self.data_seq2) >= 3:
        self.data_seq1 = hex( int( self.data_seq1, 16 ) + int( self.data_seq2[:1], 16) ).lstrip("0x")
        self.data_seq2 = self.data_seq2[-2:]
    
    if len(self.data_seq1) <= 1:
        self.data_seq1 = "0"+self.data_seq1
        
    if len(self.data_seq2) <= 1:
        self.data_seq2 = "0"+self.data_seq2
    
    
    counter = hex( self.counter ).lstrip("0x")
    if len(counter):
        counter = "0" + counter
        
    print(self.data_seq1+" "+self.data_seq2+desc)   
    data = header+self.data_seq1+self.data_seq2+self.id_com+counter+"00457874656e644f504c6f6700"
    self.zkclient.sendto(data.decode("hex"), self.address)
    print(data)
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
    except:
        bingung=1
        if self.extlogtrynumber == 1:
            self.extlogtrynumber = 2
            zkextendoplog(self)
    
    self.id_com = self.data_recv.encode("hex")[8:12]
    self.counter = self.counter+1
    print(self.data_recv.encode("hex"))
    return self.data_recv[8:]


#zkextendoplog end

#zkplatform

def zkplatform(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~Platform'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False


def zkplatformVersion(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~ZKFPVersion'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkplatform end
#zkworkcode
def zkworkcode(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = 'WorkCode'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkworkcode end

#zkssr 
def zkssr(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~SSR'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
#zkssr end

#zkpin 
def zkpinwidth(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~PIN2Width'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
#zkpin end

#zkface 
def zkfaceon(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = 'FaceFunOn'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
#zkface end

#zkserialnumber 
def zkserialnumber(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~SerialNumber'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
#zkserialnumber end

#zkdevice
def zkdevicename(self):
    """Start a connection with the time clock"""
    command = CMD_DEVICE
    command_string = '~DeviceName'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
    

def zkenabledevice(self):
    """Start a connection with the time clock"""
    command = CMD_ENABLEDEVICE
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

def zkdisabledevice(self):
    """Start a connection with the time clock"""
    command = CMD_DISABLEDEVICE
    command_string = '\x00\x00'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
#zkdevice end
#zkuser 
def getSizeUser(self):
    """Checks a returned packet to see if it returned CMD_PREPARE_DATA,
    indicating that data packets are to be sent

    Returns the amount of bytes that are going to be sent"""
    command = unpack('HHHH', self.data_recv[:8])[0] 
    if command == CMD_PREPARE_DATA:
        size = unpack('I', self.data_recv[8:12])[0]
        return size
    else:
        return False


def zksetuser(self, uid, userid, name, password, role):
    """Start a connection with the time clock"""
    command = CMD_SET_USER
    command_string = pack('sxs8s28ss7sx8s16s', chr( uid ), chr(role), password, name, chr(1), '', userid, '' )
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
    
    
def zkgetuser(self):
    """Start a connection with the time clock"""
    command = CMD_USERTEMP_RRQ
    command_string = '\x05'
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        
        
        if getSizeUser(self):
            bytes = getSizeUser(self)
            
            while bytes > 0:
                data_recv, addr = self.zkclient.recvfrom(1032)
                self.userdata.append(data_recv)
                bytes -= 1024
            
            self.session_id = unpack('HHHH', self.data_recv[:8])[2]
            data_recv = self.zkclient.recvfrom(8)
        
        users = {}
        if len(self.userdata) > 0:
            # The first 4 bytes don't seem to be related to the user
            for x in xrange(len(self.userdata)):
                if x > 0:
                    self.userdata[x] = self.userdata[x][8:]
            
            userdata = ''.join( self.userdata )
            
            userdata = userdata[11:]
            
            while len(userdata) > 72:
                
                uid, role, password, name, userid = unpack( '2s2s8s28sx31s', userdata.ljust(72)[:72] )
                
                uid = int( uid.encode("hex"), 16)
                # Clean up some messy characters from the user name
                password = password.split('\x00', 1)[0]
                password = unicode(password.strip('\x00|\x01\x10x'), errors='ignore')
                
                #uid = uid.split('\x00', 1)[0]
                userid = unicode(userid.strip('\x00|\x01\x10x'), errors='ignore')
                
                name = name.split('\x00', 1)[0]
                
                if name.strip() == "":
                    name = uid
                
                users[uid] = (userid, name, int( role.encode("hex"), 16 ), password)
                
                #print("%d, %s, %s, %s, %s" % (uid, userid, name, int( role.encode("hex"), 16 ), password))
                userdata = userdata[72:]
                
        return users
    except:
        return False
    

def zkclearuser(self):
    """Start a connection with the time clock"""
    command = CMD_CLEAR_DATA
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False


def zkclearadmin(self):
    """Start a connection with the time clock"""
    command = CMD_CLEAR_ADMIN
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkuser end
#zkattendance
def getSizeAttendance(self):
    """Checks a returned packet to see if it returned CMD_PREPARE_DATA,
    indicating that data packets are to be sent

    Returns the amount of bytes that are going to be sent"""
    command = unpack('HHHH', self.data_recv[:8])[0] 
    if command == CMD_PREPARE_DATA:
        size = unpack('I', self.data_recv[8:12])[0]
        return size
    else:
        return False


def reverseHex(hexstr):
    tmp = ''
    for i in reversed( xrange( len(hexstr)/2 ) ):
        tmp += hexstr[i*2:(i*2)+2]
    
    return tmp
    
def zkgetattendance(self):
    """Start a connection with the time clock"""
    command = CMD_ATTLOG_RRQ
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        
        if getSizeAttendance(self):
            bytes = getSizeAttendance(self)
            while bytes > 0:
                data_recv, addr = self.zkclient.recvfrom(1032)
                self.attendancedata.append(data_recv)
                bytes -= 1024
                
            self.session_id = unpack('HHHH', self.data_recv[:8])[2]
            data_recv = self.zkclient.recvfrom(8)
        
        attendance = []  
        if len(self.attendancedata) > 0:
            # The first 4 bytes don't seem to be related to the user
            for x in xrange(len(self.attendancedata)):
                if x > 0:
                    self.attendancedata[x] = self.attendancedata[x][8:]
            
            attendancedata = ''.join( self.attendancedata )
            
            attendancedata = attendancedata[14:]
            
            while len(attendancedata) > 40:
                
                uid, state, timestamp, space = unpack( '24s1s4s11s', attendancedata.ljust(40)[:40] )
                
                
                # Clean up some messy characters from the user name
                #uid = unicode(uid.strip('\x00|\x01\x10x'), errors='ignore')
                uid = uid.split('\x00', 1)[0]
                #print "%s, %s, %s" % (uid, state, decode_time( int( reverseHex( timestamp.encode('hex') ), 16 ) ) )
                
                attendance.append( ( uid, int( state.encode('hex'), 16 ), decode_time( int( reverseHex( timestamp.encode('hex') ), 16 ) ) ) )
                
                attendancedata = attendancedata[40:]
            
        return attendance
    except:
        return False
    
    
def zkclearattendance(self):
    """Start a connection with the time clock"""
    command = CMD_CLEAR_ATTLOG
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False

#zkattendane end
#zktime 
def reverseHex(hexstr):
    tmp = ''
    for i in reversed( xrange( len(hexstr)/2 ) ):
        tmp += hexstr[i*2:(i*2)+2]
    
    return tmp
    
def zksettime(self, t):
    """Start a connection with the time clock"""
    command = CMD_SET_TIME
    command_string = pack('I',encode_time(t))
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return self.data_recv[8:]
    except:
        return False
    

def zkgettime(self):
    """Start a connection with the time clock"""
    command = CMD_GET_TIME
    command_string = ''
    chksum = 0
    session_id = self.session_id
    reply_id = unpack('HHHH', self.data_recv[:8])[3]

    buf = self.createHeader(command, chksum, session_id,
        reply_id, command_string)
    self.zkclient.sendto(buf, self.address)
    #print buf.encode("hex")
    try:
        self.data_recv, addr = self.zkclient.recvfrom(1024)
        self.session_id = unpack('HHHH', self.data_recv[:8])[2]
        return decode_time( int( reverseHex( self.data_recv[8:].encode("hex") ), 16 ) )
    except:
        return False

#zktime end
#zklib
class ZKLib:
    
    def __init__(self, ip, port):
        self.address = (ip,port)
        self.zkclient = socket(AF_INET, SOCK_DGRAM)
        self.zkclient.settimeout(3)
        self.session_id = 0
        self.userdata = []
        self.attendancedata = []
    
    
    def createChkSum(self, p):
        """This function calculates the chksum of the packet to be sent to the 
        time clock

        Copied from zkemsdk.c"""
        l = len(p)
        chksum = 0
        while l > 1:
            chksum += unpack('H', pack('BB', p[0], p[1]))[0]
            
            p = p[2:]
            if chksum > USHRT_MAX:
                chksum -= USHRT_MAX
            l -= 2
        
        
        if l:
            chksum = chksum + p[-1]
            
        while chksum > USHRT_MAX:
            chksum -= USHRT_MAX
        
        chksum = ~chksum
        
        while chksum < 0:
            chksum += USHRT_MAX
        
        return pack('H', chksum)


    def createHeader(self, command, chksum, session_id, reply_id, 
                                command_string):
        """This function puts a the parts that make up a packet together and 
        packs them into a byte string"""
        buf = pack('HHHH', command, chksum,
            session_id, reply_id) + command_string
        
        buf = unpack('8B'+'%sB' % len(command_string), buf)
        
        chksum = unpack('H', self.createChkSum(buf))[0]
        #print unpack('H', self.createChkSum(buf))
        reply_id += 1
        if reply_id >= USHRT_MAX:
            reply_id -= USHRT_MAX

        buf = pack('HHHH', command, chksum, session_id, reply_id)
        return buf + command_string
    
    
    def checkValid(self, reply):
        """Checks a returned packet to see if it returned CMD_ACK_OK,
        indicating success"""
        command = unpack('HHHH', reply[:8])[0]
        if command == CMD_ACK_OK:
            return True
        else:
            return False
            
    def connect(self):
        return zkconnect(self)
            
    def disconnect(self):
        return zkdisconnect(self)
        
    def version(self):
        return zkversion(self)
        
    def osversion(self):
        return zkos(self)
        
    def extendFormat(self):
        return zkextendfmt(self)
    
    def extendOPLog(self, index=0):
        return zkextendoplog(self, index)
    
    def platform(self):
        return zkplatform(self)
    
    def fmVersion(self):
        return zkplatformVersion(self)
        
    def workCode(self):
        return zkworkcode(self)
        
    def ssr(self):
        return zkssr(self)
    
    def pinWidth(self):
        return zkpinwidth(self)
    
    def faceFunctionOn(self):
        return zkfaceon(self)
    
    def serialNumber(self):
        return zkserialnumber(self)
    
    def deviceName(self):
        return zkdevicename(self)
        
    def disableDevice(self):
        return zkdisabledevice(self)
    
    def enableDevice(self):
        return zkenabledevice(self)
        
    def getUser(self):
        return zkgetuser(self)
        
    def setUser(self, uid, userid, name, password, role):
        return zksetuser(self, uid, userid, name, password, role)
        
    def clearUser(self):
        return zkclearuser(self)
    
    def clearAdmin(self):
        return zkclearadmin(self)
        
    def getAttendance(self):
        return zkgetattendance(self)
    
    def clearAttendance(self):
        return zkclearattendance(self)
        
    def setTime(self, t):
        return zksettime(self, t)
    
    def getTime(self):
        return zkgettime(self)

#zklib end

#print(zk.getAttendance())



class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = "Test4"
    _svc_display_name_ = "Test4"
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def getItem(self,item):
        return item[2]

    def filterData(self,data):
        # with open('E:\\Log\\data.txt','a') as f:
        #     f.write(str(tuple(data)))
        #     f.writeline('-------------')
        return data[2].strftime('%Y-%m-%d')>=datetime.today().strftime('%Y-%m-%d')

    def getIsoDate(self,data):
        data=list(data)
        data[2]=data[2].isoformat()
        data[1]=data[1]
        data[0]=data[0]
        return tuple(data)

    def SvcDoRun(self):
        rc = None
        while rc != win32event.WAIT_OBJECT_0:
            try:
                zk=ZKLib('192.168.0.75',4370)
                ret=zk.connect()
                connection=client.HTTPConnection('merahunar.talbrum.com')
                attendancedata=zk.getAttendance()
                with open('E:\\Log\\data.txt','a') as f:
                    f.write(str(attendancedata))
                data=filter(self.filterData,attendancedata)
                data=list(map(self.getIsoDate,data))
                connection.request('POST','/api/v2/api_data_send/getDataService',data.__str__())
            except Exception as e:
                pass
                with open('E:\\Log\\data.txt','a') as f:
                    f.write(str(e))
            rc = win32event.WaitForSingleObject(self.hWaitStop, 15000)
    sleep(5)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)
