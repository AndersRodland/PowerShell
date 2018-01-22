$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment

#The chassis is the physical container that houses the components of a computer. Check if the machine’s chasis type is 9.Laptop 10.Notebook 14.Sub-Notebook
if (Get-WmiObject -Class win32_systemenclosure -ComputerName $computer | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14}) { $tsenv.Value("IsLaptop") = "TRUE" }

#Shows battery status , if true then the machine is a laptop.
if (Get-WmiObject -Class win32_battery -ComputerName $computer) { $tsenv.Value("IsLaptop") = "TRUE" }
