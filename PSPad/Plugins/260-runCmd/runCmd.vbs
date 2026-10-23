'Allows you run an external cmd.exe command by setting the cmdArgs variable as desired

const module_name	= "runCMD"
const module_ver		= "1.0"
const module_title	= "runCMD"

sub Init
  addMenuItem "&RunCMD", "", "externalCommand"
end sub

sub externalCommand

  cmdArgs = "notepad.exe" 
   'NOTE: cmdArgs can contain multiple commands by separating them with && like this: cmdArgs = "cd\php && php.exe"
   
  Set wshShell = CreateObject( "WScript.Shell" )

  wshShell.Run "cmd.exe /K " & cmdArgs, 1, False   
   'NOTE:                               1 = show dos window
   '                                    0 = hide dos window
   '                    /K = keep dos window open when application terminates
   '                    /C = close dos window when application terminates  
               
  Set wshShell = Nothing

end sub