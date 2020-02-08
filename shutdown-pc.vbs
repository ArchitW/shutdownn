Call WaitRoutine() 

Sub WaitRoutine()

                Dim intWaitedTime
                Dim intMinsToWait : intMinsToWait = -1
  Dim objshell


   set objShell = CreateObject("WScript.Shell") 
                strAnswer = InputBox("How many minutes do you wish to wait?","Shutdown Computer")
                on error resume next
                intMinsToWait = cint(strAnswer)
                if (intMinsToWait = -1) or (strAnswer = "") then
                                msgbox "Shutdown Cancelled.",vbokonly + vbexclamation,"Shutdown Computer"
                                exit sub
                end if               

                NewDate = DateAdd("N", intMinsToWait, now())
                If msgbox("This will shutdown the computer at " & NewDate & ". Continue?", vbquestion + vbyesno,"Shutdown Computer") = vbno then
   exit sub   
                End If               

                'make script sleep
                WScript.Sleep(intMinsToWait * 60 * 1000) 

  strShutdown = "shutdown -s -t 0 -f -m \\" & "."
  objShell.Run strShutdown
  Wscript.Quit
End sub
