import win32com.client, pythoncom, time, msvcrt

class ServerEvents:
    def OnServerShutDown(self, Reason):
        #Works!
        print ('Server Shutdown at', time.ctime())

class GroupsEvents:
    def OnGlobalDataChange(self, TransactionID, GroupHandle, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps):
        pass
        #Works!
        #print 'Global DataChange, transaction ID', TransactionID
        #for i in range(NumItems):
        #   print ItemValues[i],Qualities[i],TimeStamps[i]

class GroupEvents:
    def OnDataChange(self, TransactionID, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps):
        #Works!
        print ("Group 1 data changed at", time.ctime())
        for i in range(NumItems):
            print (TransactionID,ItemValues[i],Qualities[i],TimeStamps[i])

    def OnAsyncReadComplete(self, TransactionID, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps, Errors):
        print ("AsyncRead Returned!")
        print (ItemValues)
        for i in range(NumItems):
            print (ItemValues[i],Qualities[i],TimeStamps[i])

    def OnAsyncWriteComplete(self, TransactionID, NumItems, ClientHandles, Errors):
        print ("AsyncWrite No. %2d Completed"%(TransactionID))
        print ("Errors:", Errors)
        print ("Client Handles:", ClientHandles)

    def OnAsyncCancelComplete(self, CancelID):
        #Not working, not in VB either
        print ("Async Request Canceled", CancelID)

class Group2Events:
    def OnDataChange(self, TransactionID, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps):
        #Works!
        print ("Group 2 data changed at", time.ctime())
        for i in range(NumItems):
            print (TransactionID,ItemValues[i],Qualities[i],TimeStamps[i])

    def OnAsyncReadComplete(self, TransactionID, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps, Errors):
        print ("AsyncRead Returned!")
        print (ItemValues)
        for i in range(NumItems):
            print (ItemValues[i],Qualities[i],TimeStamps[i])

    def OnAsyncWriteComplete(self, TransactionID, NumItems, ClientHandles, Errors):
        print ("AsyncWrite No. %2d Completed"%(TransactionID))
        print ("Errors:", Errors)
        print ("Client Handles:", ClientHandles)

    def OnAsyncCancelComplete(self, CancelID):
        #Not working, not in VB either
        print ("Async Request Canceled", CancelID)

print ("Starting======>")
OPC=win32com.client.DispatchWithEvents('OPC.Automation.1',ServerEvents)

try:
    OPC.Connect('KEPware.KEPServerEx.V4')
except:
    print("Something wrong with connect")
    print("Try another OPC Server")
    raise SystemExit
    
groups=OPC.OPCGroups
groups_events=win32com.client.WithEvents(groups,GroupsEvents)
group=groups.Add('Group 1')
group.UpdateRate=10000
group.IsSubscribed=1
group.IsActive=1
group_events=win32com.client.WithEvents(group,GroupEvents)
item=group.OPCItems.AddItem('Channel_1.Device_1.Tag_1', 1)
item2=group.OPCItems.AddItem('Channel_0_User_Defined.Sine.Sine1', 2)

group2=groups.Add('Group 2')
group2.UpdateRate=1000
group2.IsSubscribed=1
group2.IsActive=1
group2_events=win32com.client.WithEvents(group2,Group2Events)
item3=group2.OPCItems.AddItem('Channel_1.Device_1.Tag_2', 21)
item4=group2.OPCItems.AddItem('Channel_2.Device_3.Tag_1', 22)

#OPC_Device=0x1
#sh=(0,item.ServerHandle,item2.ServerHandle)
#err=None
#retval=None
#retval=group.AsyncRead(2,sh,err,123)
#Works if you pad the ServerHandle array with an initial 0
#retval=group.AsyncRefresh(OPC_Device,123)
#Works!!!
#ret=group.AsyncCancel(retval)
#Works!!!
#retval=group.AsyncWrite(2,sh,(0,.5,2),err,123)
#Works if you pad ServerHandle and Value arrays with an initial 0
#retval=item.Read(OPC_Device)
#Works!!!
#print retval

end = time.clock() + 60
while (time.clock() < end and not msvcrt.kbhit()):
    if pythoncom.PumpWaitingMessages(): #@UndefinedVariable
        break # maybe broken
        

OPC.Disconnect()    
print ("Done") 