'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' Copy\move file
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function rocheck(filik)
  Dim ROdis
  If objfso.FileExists(filik) Then
    Set ROdis = objFSO.GetFile(filik)
    If ROdis.Attributes And ReadOnly Then
      rocheck = true
    else
      rocheck = false
    End if
  else
    rocheck = false
  end if
end function

function roremove(filik)
  Dim ROdis
  If objfso.FileExists(filik) Then
    Set ROdis = objFSO.GetFile(filik)
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Action ]" & logstartline() & "Marking " & filik & " as writable...", 1, ForAppending, 0)
    On Error Resume Next
    ROdis.Attributes = ROdis.Attributes XOR ReadOnly
    If Err.Number 0 Then
      errflg = 1
      call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Error ]" & logstartline() & filik & " failed to be marked as writable with error #" & CStr(Err.Number) & " " & Err.Description & ". Source: " & Err.Source, "Failed to make writable " & objFSO.GetFileName(source) & "!", ForAppending, 3)
      On Error Goto 0
      roremove = false
    else
      call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[OK ]" & logstartline() & filik & " marked as writable", 1, ForAppending, 0)
      roremove = true
    end if
    On Error Goto 0
  else
    roremove = false
  end if
end function

Function filemover(source, destination, movflg, overflg)
  Dim overwritebool
  Dim destro
  source = Replace(source, "\\", "\", 1, 1)
  if Left(source,1) = "\" Then
    source = "\" & source
  end if
  destination = Replace(destination, "\\", "\", 1, 1)
  if Left(destination,1) = "\" Then
    destination = "\" & destination
  end if
  If overflg = 1 Then
    overwritebool = true
  Else
    overwritebool = false
  End if
  errflg = 0
  If objfso.FileExists(source) = False Then
    errflg = 1
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Warning]" & logstartline() & "" & source & " does not exist!", 1, ForAppending, 2)
    Exit function
  End if
  if movflg = 1 AND rocheck(source) = true Then
    if roremove(source) = false Then
      Exit Function
    end if
  End if
  If objfso.FolderExists(destination) Then
    if Right(destination, 1) "\" Then
      destination = destination & "\" & objFSO.GetFileName(source)
    else
      destination = destination & objFSO.GetFileName(source)
    End if
  else
    if objfso.FolderExists(destination & "\") Then
      destination = destination & "\" & objFSO.GetFileName(source)
    else
      If objfso.FileExists(destination) = False Then
        dim flname
        flname = objfso.GetFileName(destination)
        if objfso.FolderExists(Replace(destination, flname, "")) = False Then
          call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Error ]" & logstartline() & "Destination folder " & Replace(destination, flname, "") & " does not exist!", "1", ForAppending, 3)
          Exit function
        end if
      end if
    end if
  end if
  destro = rocheck(destination)
  If objfso.FileExists(destination) AND destro = true Then
    if roremove(destination) = false Then
      Exit Function
    end if
  end if
  On Error Resume Next
  call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Action ]" & logstartline() & "Copying\moving " & source & "to " & destination & "...", 1, ForAppending, 0)
  objFSO.CopyFile source, destination, overwritebool
  if movflg = 1 Then
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Action ]" & logstartline() & "Removing " & source & "...", 1, ForAppending, 0)
    objFSO.DeleteFile source, true
  end if
  if destro = true Then
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Action ]" & logstartline() & "Restoring read-only attribute for " & destination & "...", 1, ForAppending, 0)
    Set rodis = objFSO.GetFile(destination)
    rodis.Attributes = rodis.Attributes + ReadOnly
    If Err.Number 0 Then
      call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Error ]" & logstartline() & "Failed to restore read-only attribute for " & destination & " with error #" & CStr(Err.Number) & " " & Err.Description & ". Source: " & Err.Source, "Failed to make writable " & objFSO.GetFileName(source) & "!", ForAppending, 3)
    end if
  end if
  If Err.Number 0 Then
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[Error ]" & logstartline() & source & " failed to be copied\moved to " & destination & " with error #" & CStr(Err.Number) & " " & Err.Description & ". Source: " & Err.Source, "Failed to copy\move " & objFSO.GetFileName(source) & "!", ForAppending, 3)
  else
    call writeinfile (HTA_Log & dateyyymmdd() & ".log", "[OK ]" & logstartline() & source & " copied/moved to " & destination, 1, ForAppending, 0)
  end if
  On Error Goto 0
End Function
