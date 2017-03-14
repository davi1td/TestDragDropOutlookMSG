'
' Created by SharpDevelop.
' User: davi1td
' Date: 11/13/2016
' Time: 10:06 AM
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Imports Excel = NetOffice.ExcelApi
Class ipFunctions
'    Copyright 2010-2016 Thomas Rohmer-Kretz

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'    http://trk.free.fr/ipcalc/

'    Visual Basic for Excel

'==============================================
'   IP v4
'==============================================

'----------------------------------------------
'   IpIsValid
'----------------------------------------------
' Returns true if an ip address is formated exactly as it should be:
' no space, no extra zero, no incorrect value
Function IpIsValid(ByVal ip As String) As Boolean
    IpIsValid = (IpBinToStr(IpStrToBin(ip)) = ip)
End Function

'----------------------------------------------
'   IpStrToBin
'----------------------------------------------
' Converts a text IP address to binary
' example:
'   IpStrToBin("1.2.3.4") returns 16909060
Function IpStrToBin(ByVal ip As String) As Double
    Dim pos As Integer
    ip = ip + "."
    IpStrToBin = 0

    While ip <> ""
        pos = InStr(ip, ".")
        IpStrToBin = IpStrToBin * 256 + Val(Left(ip, pos - 1))
        ip = Mid(ip, pos + 1)
    End While 
End Function

'----------------------------------------------
'   IpBinToStr
'----------------------------------------------
' Converts a binary IP address to text
' example:
'   IpBinToStr(16909060) returns "1.2.3.4"
Function IpBinToStr(ByVal ip As Double) As String
    Dim divEnt As Double
    Dim i As Integer
    i = 0
    IpBinToStr = ""
    While i < 4
        If IpBinToStr <> "" Then IpBinToStr = "." + IpBinToStr
        divEnt = Int(ip / 256)
        IpBinToStr = Format(ip - (divEnt * 256)) + IpBinToStr
        ip = divEnt
        i = i + 1
    End While
End Function

'----------------------------------------------
'   IpAdd
'----------------------------------------------
' example:
'   IpAdd("192.168.1.1"; 4) returns "192.168.1.5"
'   IpAdd("192.168.1.1"; 256) returns "192.168.2.1"
Function IpAdd(ByVal ip As String, offset As Double) As String
    IpAdd = IpBinToStr(IpStrToBin(ip) + offset)
End Function

'----------------------------------------------
'   IpAnd
'----------------------------------------------
' logical AND
' example:
'   IpAnd("192.168.1.1"; "255.255.255.0") returns "192.168.1.0"
Function IpAnd(ByVal ip1 As String, ByVal ip2 As String) As String
    ' compute logical AND from right to left
    Dim result As String
    While ((ip1 <> "") And (ip2 <> ""))
        Call IpBuild(IpParse(ip1) And IpParse(ip2), result)
    End While
    IpAnd = result
End Function

'----------------------------------------------
'   IpAdd2
'----------------------------------------------
' another implementation of IpAdd which not use the binary representation
Function IpAdd2(ByVal ip As String, offset As Double) As String
    Dim result As String
    While (ip <> "")
        offset = IpBuild(IpParse(ip) + offset, result)
   End While
    IpAdd2 = result
End Function

'----------------------------------------------
'   IpGetByte
'----------------------------------------------
' get one byte from an ip address given its position
' example:
'   IpGetByte("192.168.1.1"; 1) returns 192
Function IpGetByte(ByVal ip As String, pos As Integer) As Integer
    pos = 4 - pos
    For i = 0 To pos
        IpGetByte = IpParse(ip)
    Next
End Function

'----------------------------------------------
'   IpSetByte
'----------------------------------------------
' set one byte in an ip address given its position and value
' example:
'   IpSetByte("192.168.1.1"; 4; 20) returns "192.168.1.20"
Function IpSetByte(ByVal ip As String, pos As Integer, newvalue As Integer) As String
    Dim result As String
    Dim byteval As Double
    Dim I As Integer 
    i = 4
    While (ip <> "")
        byteval = IpParse(ip)
        If (i = pos) Then byteval = newvalue
        Call IpBuild(byteval, result)
        i = i - 1
    End While
    IpSetByte = result
End Function

'----------------------------------------------
'   IpMask
'----------------------------------------------
' returns an IP netmask from a subnet
' both notations are accepted
' example:
'   IpMask("192.168.1.1/24") returns "255.255.255.0"
'   IpMask("192.168.1.1 255.255.255.0") returns "255.255.255.0"
Function IpMask(ByVal ip As String) As String
    IpMask = IpBinToStr(IpMaskBin(ip))
End Function

'----------------------------------------------
'   IpWildMask
'----------------------------------------------
' returns an IP Wildcard (inverse) mask from a subnet
' both notations are accepted
' example:
'   IpWildMask("192.168.1.1/24") returns "0.0.0.255"
'   IpWildMask("192.168.1.1 255.255.255.0") returns "0.0.0.255"
Function IpWildMask(ByVal ip As String) As String
    IpWildMask = IpBinToStr(((2 ^ 32) - 1) - IpMaskBin(ip))
End Function

'----------------------------------------------
'   IpInvertMask
'----------------------------------------------
' returns an IP Wildcard (inverse) mask from a subnet mask
' or a subnet mask from a wildcard mask
' example:
'   IpWildMask("255.255.255.0") returns "0.0.0.255"
'   IpWildMask("0.0.0.255") returns "255.255.255.0"
Function IpInvertMask(ByVal mask As String) As String
    IpInvertMask = IpBinToStr(((2 ^ 32) - 1) - IpStrToBin(mask))
End Function

'----------------------------------------------
'   IpMaskLen
'----------------------------------------------
' returns prefix length from a mask given by a string notation (xx.xx.xx.xx)
' example:
'   IpMaskLen("255.255.255.0") returns 24 which is the number of bits of the subnetwork prefix
Function IpMaskLen(ByVal ipmaskstr As String) As Integer
	Dim notMask As Double
	Dim zeroBits As Integer
    notMask = 2 ^ 32 - 1 - IpStrToBin(ipmaskstr)
    zeroBits = 0
    Do While notMask <> 0
        notMask = Int(notMask / 2)
        zeroBits = zeroBits + 1
    Loop
    IpMaskLen = 32 - zeroBits
End Function

'----------------------------------------------
'   IpWithoutMask
'----------------------------------------------
' removes the netmask notation at the end of the IP
' example:
'   IpWithoutMask("192.168.1.1/24") returns "192.168.1.1"
'   IpWithoutMask("192.168.1.1 255.255.255.0") returns "192.168.1.1"
Function IpWithoutMask(ByVal ip As String) As String
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
    End If
    If (p = 0) Then
        IpWithoutMask = ip
    Else
        IpWithoutMask = Left(ip, p - 1)
    End If
End Function

'----------------------------------------------
'   IpSubnetLen
'----------------------------------------------
' get the mask len from a subnet
' example:
'   IpSubnetLen("192.168.1.1/24") returns 24
'   IpSubnetLen("192.168.1.1 255.255.255.0") returns 24
Function IpSubnetLen(ByVal ip As String) As Integer
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
        If (p = 0) Then
            IpSubnetLen = 32
        Else
            IpSubnetLen = IpMaskLen(Mid(ip, p + 1))
        End If
    Else
        IpSubnetLen = Val(Mid(ip, p + 1))
    End If
End Function

'----------------------------------------------
'   IpIsInSubnet
'----------------------------------------------
' returns TRUE if "ip" is in "subnet"
' subnet must have the / mask notation (xx.xx.xx.xx/yy)
' example:
'   IpIsInSubnet("192.168.1.35"; "192.168.1.32/29") returns TRUE
'   IpIsInSubnet("192.168.1.35"; "192.168.1.32 255.255.255.248") returns TRUE
'   IpIsInSubnet("192.168.1.41"; "192.168.1.32/29") returns FALSE
Function IpIsInSubnet(ByVal ip As String, ByVal subnet As String) As Boolean
    'IpIsInSubnet = (IpAnd(ip, IpMask(subnet)) = IpWithoutMask(subnet))
    ' the following line also works with non standard subnet notation:
    IpIsInSubnet = (IpAnd(ip, IpMask(subnet)) = (IpAnd(IpWithoutMask(subnet), IpMask(subnet))))
End Function
function IpSubnetIsInSubnet(ByVal subnet1 As String, ByVal subnet2 As String) As Boolean
    Dim Mask1 As Double
    Dim Mask2 As Double
    Mask1 = IpMaskBin(subnet1)
    Mask2 = IpMaskBin(subnet2)
    If (Mask1 < Mask2) Then
        IpSubnetIsInSubnet = False
    ElseIf IpIsInSubnet(IpWithoutMask(subnet1), subnet2) Then
        IpSubnetIsInSubnet = True
    Else
        IpSubnetIsInSubnet = False
    End If
End Function

'----------------------------------------------
'   IpSubnetSize
'----------------------------------------------
' returns the number of addresses in a subnet
' example:
'   IpSubnetSize("192.168.1.32/29") returns 8
'   IpSubnetSize("192.168.1.0 255.255.255.0") returns 256
Function IpSubnetSize(ByVal subnet As String) As Double
    IpSubnetSize = 2 ^ (32 - IpSubnetLen(subnet))
End Function

'----------------------------------------------
'   IpSubnetInSubnetVLookup
'----------------------------------------------
' tries to match a subnet against a list of subnets in the left-most
' column of table_array and returns the value in the same row based on the
' index_number
' the value matches if 'subnet' is equal or included in one of the subnets
' in the array
' 'subnet' is the value to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' 'index_number' is the column number in table_array from which the matching
'      value must be returned. The first column which contains subnets is 1.
' note: add the subnet 0.0.0.0/0 at the end of the array if you want the
' function to return a default value
Function IpSubnetInSubnetVLookup(ByVal subnet As String, table_array As Excel.Range, index_number As Integer) As String
    IpSubnetInSubnetVLookup = "Not Found"
    For i = 1 To table_array.Rows.Count
        If IpSubnetIsInSubnet(subnet, table_array.Cells(i, 1).Value ) Then
            IpSubnetInSubnetVLookup = table_array.Cells(i, index_number).value
            Exit For
        End If
    Next i
End Function

'----------------------------------------------
'   IpSubnetInSubnetMatch
'----------------------------------------------
' tries to match a subnet against a list of subnets in the left-most
' column of table_array and returns the row number
' the value matches if 'subnet' is equal or included in one of the subnets
' in the array
' 'subnet' is the value to search for in the subnets in the first column of
'      the table_array
' 'table_array' is one or more columns of data
' returns 0 if the subnet is not included in any of the subnets from the list
Function IpSubnetInSubnetMatch(ByVal subnet As String, table_array As Excel.Range) As Integer
    IpSubnetInSubnetMatch = 0
    For i = 1 To table_array.Rows.Count
        If IpSubnetIsInSubnet(subnet, table_array.Cells(i, 1).value) Then
            IpSubnetInSubnetMatch = i
            Exit For
        End If
    Next i
End Function

'----------------------------------------------
'   IpIsPrivate
'----------------------------------------------
' returns TRUE if "ip" is in one of the private IP address ranges
' example:
'   IpIsPrivate("192.168.1.35") returns TRUE
'   IpIsPrivate("209.85.148.104") returns FALSE
Function IpIsPrivate(ByVal ip As String) As Boolean
    IpIsPrivate = (IpIsInSubnet(ip, "10.0.0.0/8") Or IpIsInSubnet(ip, "172.16.0.0/12") Or IpIsInSubnet(ip, "192.168.0.0/16"))
End Function

'----------------------------------------------
'   IpDiff
'----------------------------------------------
' difference between 2 IP addresses
' example:
'   IpDiff("192.168.1.7"; "192.168.1.1") returns 6
Function IpDiff(ByVal ip1 As String, ByVal ip2 As String) As Double
    Dim mult As Double
    mult = 1
    IpDiff = 0
    While ((ip1 <> "") Or (ip2 <> ""))
        IpDiff = IpDiff + mult * (IpParse(ip1) - IpParse(ip2))
        mult = mult * 256
    End While
End Function

'----------------------------------------------
'   IpParse
'----------------------------------------------
' Parses an IP address by iteration from right to left
' Removes one byte from the right of "ip" and returns it as an integer
' example:
'   if ip="192.168.1.32"
'   IpParse(ip) returns 32 and ip="192.168.1" when the function returns
Function IpParse(ByRef ip As String) As Integer
    Dim pos As Integer
    pos = InStrRev(ip, ".")
    If pos = 0 Then
        IpParse = Val(ip)
        ip = ""
    Else
        IpParse = Val(Mid(ip, pos + 1))
        ip = Left(ip, pos - 1)
    End If
End Function

'----------------------------------------------
'   IpBuild
'----------------------------------------------
' Builds an IP address by iteration from right to left
' Adds "ip_byte" to the left the "ip"
' If "ip_byte" is greater than 255, only the lower 8 bits are added to "ip"
' and the remaining bits are returned to be used on the next IpBuild call
' example 1:
'   if ip="168.1.1"
'   IpBuild(192, ip) returns 0 and ip="192.168.1.1"
' example 2:
'   if ip="1"
'   IpBuild(258, ip) returns 1 and ip="2.1"
Function IpBuild(ip_byte As Double, ByRef ip As String) As Double
    If ip <> "" Then ip = "." + ip
    ip = Format(ip_byte And 255) + ip
    IpBuild = ip_byte \ 256
End Function

'----------------------------------------------
'   IpMaskBin
'----------------------------------------------
' returns binary IP mask from an address with / notation (xx.xx.xx.xx/yy)
' example:
'   IpMask("192.168.1.1/24") returns 4294967040 which is the binary
'   representation of "255.255.255.0"
Function IpMaskBin(ByVal ip As String) As Double
    Dim bits As Integer
    bits = IpSubnetLen(ip)
    IpMaskBin = (2 ^ bits - 1) * 2 ^ (32 - bits)
End Function

Function Hex2Bin(ByVal strHex As String) As Double
    Dim v As Double
    For i = 1 To Len(strHex)
        v = 16 * v + Val("&H" & Mid$(strHex, i, 1))
    Next
    Hex2Bin = v
End Function
	
End Class 
