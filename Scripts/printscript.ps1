#$STAFFPATH    =
#CCWB
#CCTB
#WB
#TB
#FM

#LF
#VF
#PH
#FR

#STAMS
#STARS

Start-Process -FilePath "C:\Users\Kudo\Documents\Life\Resumes\LaTeX Resume latest.pdf" -Verb Print -PassThru | %{sleep 10;$_} | kill
