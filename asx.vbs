dim linkedrecord(2)
linkedrecord(0) = "CONVERSION_TEST_60_CUS_A1,CONVERSION_TEST_60,CUS,customerdetails,custId,SIN_70,INS,1,A1"
linkedrecord(1) = "CONVERSION_TEST_60_CUS_A2,CONVERSION_TEST_60,CUS,linkedcrn,CrnID,CONVERSION_TEST_60_CUS,INS,1,A2"

'generate data function
sub generatedata(inserts(), table, start_num, end_num)
    dim outfile
    dim fso

    Set fso = createobject("Scripting.FileSystemObject") 
    FilePath = "ASX_" & table & ".csv"
    If fso.FileExists(FilePath) Then
        fso.DeleteFile FilePath
    end if

    set outfile = fso.CreateTextFile(FilePath, 8)
    for r = cint(start_num) to cint(end_num)
        for i = 0 to ubound(inserts)
            if len(inserts(i)) > 0 then        
               replaced = Replace(inserts(i),"TEST_60","TEST_" & cstr(r))
               outfile.Write trim(replaced) & vbCrLf
            end if
        next
    next
    outfile.Close
end sub

'run the program
set args = Wscript.Arguments.Unnamed
if args.count > 1 then
   start_num = WScript.Arguments.Item(0)
   end_num = WScript.Arguments.Item(1)
   generatedata linkedrecord, "linkedrecord", start_num, end_num
else
   Wscript.echo "ERROR! Command syntax is : cscript generate_data.vbs <start number> <end number>'"
end if
