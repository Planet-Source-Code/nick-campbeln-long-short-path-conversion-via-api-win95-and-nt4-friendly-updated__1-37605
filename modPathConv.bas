Attribute VB_Name = "modPathConv"
Option Explicit
'#####################################################################################
'#  Long <-> Short Path Conversion via API (Win95 and NT4 Friendly) (modPathConv.bas)
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.0.1 (Aug 6, 2002):
'#              Fixed a (very) stupid coding error in GetShortPath() - Dim'ed lLen and was using lRetVal - D'oh!
'#          1.0 (Aug 4, 2002):
'#              Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37605&lngWId=1
'#####################################################################################

    '#### Functions used for GetShortPath()/GetLongPath()
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long



'#####################################################################################
'# Public subs/functions
'#####################################################################################
'#########################################################
'# Converts a short path into it's long path equilvent
'#########################################################
Public Function GetLongPath(ByVal sShortPath As String) As String
    Dim lLen As Long

        '#### Setup the error handling and setup the buffer for the API call
    On Error GoTo GetLongPath_APIError
    GetLongPath = Space(1024)

        '#### Call the API, strip away the unwanted characters and return
    lLen = GetLongPathName(sShortPath, GetLongPath, Len(GetLongPath))
    GetLongPath = Left(GetLongPath, lLen)
    Exit Function

GetLongPath_APIError:
        '#### If we make it here the GetLongPathName() API does not exist, so call GetLongFromShortPath() to determine the long path
    GetLongPath = GetLongFromShortPath(sShortPath)
End Function



'#########################################################
'# Converts a long path into it's 8.5 short path equilvent
'#########################################################
Public Function GetShortPath(ByVal sLongPath As String) As String
    Dim lLen As Long

        '#### Setup the buffer for the API call
    GetShortPath = Space(1024)

        '#### Call the API, strip away the unwanted characters and return
    lLen = GetShortPathName(sLongPath, GetShortPath, Len(GetShortPath))
    GetShortPath = Left(GetShortPath, lLen)
End Function


'#####################################################################################
'# Private subs/functions
'#####################################################################################
'#########################################################
'# Converts a short path into it's long path equilvent via the Win95/NT4 supported GetShortPathName() API
'#########################################################
Private Function GetLongFromShortPath(ByVal sShortPath As String) As String
    Dim sShortElement As String
    Dim sLongElement As String
    Dim sSearchPath As String
    Dim lTilde As Long
    Dim lPathSep As Long
    Dim lIndex As Long

        '#### Fix sShortPath, default the return value, init sShortPath and default lIndex
    GetLongFromShortPath = FixPath(sShortPath)
    sShortPath = GetShortPath(GetLongFromShortPath)
    lIndex = 1

        '#### If GetShortPath() returned a value, the path in sShortPath is valid
    If (Len(sShortPath) > 0) Then
            '#### Determine the index of the first lTilde
        lTilde = InStr(1, sShortPath, "~", vbBinaryCompare)

            '#### As long as we keep finding new lTilde's in sShortPath
        Do While (lTilde > 0)
                '#### Find the first "\" preceding lTilde
            lPathSep = InStrRev(sShortPath, "\", lTilde, vbBinaryCompare)

                '#### If a lPathSep was found preceding lTilde
            If (lPathSep > 0) Then
                    '#### Peal off the path portion preceding lTilde
                sSearchPath = Left(sShortPath, lPathSep)

                '#### Else there is not a lPathSep preceding lTilde
            Else
                    '#### Default to the current directory
                sSearchPath = ".\"
            End If

                '#### Find the first "\" following lTilde
            lIndex = InStr(lTilde + 1, sShortPath, "\", vbBinaryCompare)

                '#### If a "\" was not found following lTilde, set lIndex to the Len() of sShortPath
            If (lIndex = 0) Then lIndex = Len(sShortPath)

                '#### Peal out the sShortElement we're looking for
            sShortElement = Mid(sShortPath, lPathSep + 1, lIndex - lPathSep - 1)

                '#### If this element is a directory
            If (Len(Dir(sSearchPath & sShortElement, vbDirectory)) > 0) Then
                    '#### Search for any directories starting with the same characters as sShortElement, setting sLongElement with the first value
                sLongElement = Dir(sSearchPath & Left(sShortElement, 1) & "*", vbDirectory)

                '#### Else it must be a file
            Else
                    '#### Search for any elements starting with the same characters as sShortElement, setting sLongElement with the first value
                sLongElement = Dir(sSearchPath & Left(sShortElement, 1) & "*")
            End If

                '#### While sLongElement has a value
            Do While (Len(sLongElement) > 0)
'!' Possibially make this compairson more efficient?
                    '#### If the short path matches the converted long path
                If (LCase(sSearchPath & sShortElement) = LCase(GetShortPath(sSearchPath & sLongElement))) Then
                        '#### Replace the sShortElement with sLongElement in the return value and exit the loop
                    GetLongFromShortPath = Replace(GetLongFromShortPath, sShortElement, sLongElement, , 1, vbTextCompare)
                    Exit Do

                    '#### Else we need to keep looking for a matching path
                Else
                        '#### Set sLongElement for the next loop (NOTE: If the path is not found does not mean it's an error, it simply means that the path has a tilde in it's name, ie - "C:\~Temp\Some Long Path\")
                    sLongElement = Dir()
                End If
            Loop

                '#### Determine the index of the next lTilde
            lTilde = InStr(lIndex, sShortPath, "~", vbBinaryCompare)
        Loop

        '#### Else the path in sShortPath is invalid, so return ""
    Else
        GetLongFromShortPath = ""
    End If
End Function


'#########################################################
'# Fixes the passed sPath, setting the proper path seperators in their proper positions
'#########################################################
Private Function FixPath(ByVal sPath As String) As String
        '#### Trim and fix the sPath
    FixPath = Replace(Trim(sPath), "/", "\")

        '#### Replace any non-leading double "\\"s with a single "\"
    If (Left(FixPath, 2) = "\\") Then
        FixPath = "\\" & Replace(Right(FixPath, Len(FixPath) - 2), "\\", "\")
    Else
        FixPath = Replace(FixPath, "\\", "\")
    End If

        '#### If the path ends in a slash, peal it off before we return
    If (Right(FixPath, 1) = "\") Then
        FixPath = Left(FixPath, Len(FixPath) - 1)
    End If
End Function
