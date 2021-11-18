
-------Split CSV into different files of 500000 row increments with the name splitfile


c:\Users\spanek\Desktop



$i=0; Get-Content c:\Users\spanek\Desktop\ALFRED_1.csv -ReadCount 500000 | %{$i++; $_ | Out-File c:\Users\spanek\Desktop\splitfile2_$i.csv}
