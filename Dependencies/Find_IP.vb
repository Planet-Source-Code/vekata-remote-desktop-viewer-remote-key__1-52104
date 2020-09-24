'****************************************************************
'  File:    getipname.vbs  (WSH for VBscript)
'  Author:           M. Gallant    09/30/2001
'
'  Based on script by M. Harris & T. Lavedas:
'  posted to: microsoft.public.scripting.vbscript  2000/07/21  

'  Reads IP addresses via:
'    ipconfig.exe  (NT4 and Win2000)
'    winipcfg.exe  (Win95)
'  For NT4, Win2000 resolves IP addresses to FQDN names via:
'     nslookup.exe     (with default DNS server)
' 
'****************************************************************
'Modified 2/2/02 to just show the IP addresses - Doug Knox
'Original script located at:  http://home.istar.ca/~neutron/wsh/IPInfo/getipname.html
arAddresses = GetIPAddresses()

info = ""

for each ip in arAddresses
    info = info & ip & vbTab & GetFQDN(ip) & vbCR
next

  WScript.echo info

Function GetFQDN(ipaddress)
'====
' Returns Fully Qualified Domain Name
' from reverse DNS lookup via nslookup.exe
' only implemented for NT4, 2000
'====
  set sh = createobject("wscript.shell")
  set fso = createobject("scripting.filesystemobject")
  Set Env = sh.Environment("PROCESS")

  if Env("OS") = "Windows_NT" then
    workfile = fso.gettempname
    sh.run "%comspec% /c nslookup " & ipaddress & "  > " & workfile,0,true
   set sh = nothing
   set ts = fso.opentextfile(workfile)
   data = split(ts.readall,vbcr)
   ts.close
   set ts = nothing
   fso.deletefile workfile
   set fso = nothing
  for n = 0 to ubound(data)
    if instr(data(n),"Name") then
      parts = split(data(n),":")
        hostname= trim(cstr(parts(1)))
       Exit For
    end if
    hostname = "could not resolve IP address"
  next
    GetFQDN = hostname
  else
   set sh = nothing
   set fso = nothing
   GetFQDN = ""
  end if
End Function


Function GetIPAddresses()
'=====
' Returns array of IP Addresses as output
' by ipconfig or winipcfg...
'
' Win98/WinNT have ipconfig (Win95 doesn't)
' Win98/Win95 have winipcfg (WinNt doesn't)
'
' Note: The PPP Adapter (Dial Up Adapter) is
' excluded if not connected (IP address will be 0.0.0.0)
' and included if it is connected.
'=====
  set sh = createobject("wscript.shell")
  set fso = createobject("scripting.filesystemobject")

  Set Env = sh.Environment("PROCESS")
  if Env("OS") = "Windows_NT" then
    workfile = fso.gettempname
    sh.run "%comspec% /c ipconfig > " & workfile,0,true
  else
    'winipcfg in batch mode sends output to
    'filename winipcfg.out
    workfile = "winipcfg.out"
    sh.run "winipcfg /batch" ,0,true
  end if
  set sh = nothing
  set ts = fso.opentextfile(workfile)
  data = split(ts.readall,vbcr)
  ts.close
  set ts = nothing
  fso.deletefile workfile
  set fso = nothing
  arIPAddress = array()
  index = -1
  for n = 0 to ubound(data)
    if instr(data(n),"IP Address") then
      parts = split(data(n),":")
      if trim(parts(1)) <> "0.0.0.0" then
        index = index + 1
        ReDim Preserve arIPAddress(index)
        arIPAddress(index)= trim(cstr(parts(1)))
      end if
    end if
  next
  GetIPAddresses = arIPAddress
End Function

