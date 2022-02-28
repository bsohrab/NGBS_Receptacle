import os
import win32com.client as win32
import tkinter as tk
import glob
import openpyxl

def mandatoryextraction(value, row, worksheet):
    header = []
    while value <= 2500:  # cellrow < 2500:
        criteria = ""
        cell = "$W$" + str(value)

        RorFcheck = [worksheet["$P" + cell[2:]].value, worksheet["$Q" + cell[2:]].value]
        if RorFcheck[0] is None or "X" in str(RorFcheck[0]) or RorFcheck[1] is None or "X" in str(RorFcheck[1]):
            RorFcheck = None
        elif RorFcheck[0] == True and RorFcheck[1] == False:
            RorFcheck = "Rough"
        elif RorFcheck[0] == False and RorFcheck[1] == True:
            RorFcheck = "Final"
        elif RorFcheck[0] == True and RorFcheck[1] == True:
            RorFcheck = "Rough/Final"
        elif RorFcheck[0] == False and RorFcheck[1] == False:
            RorFcheck = "Error Error Error"

        ##

        mandatoryitem = worksheet["$R" + cell[2:]].value
        practicetext = worksheet["$L" + cell[2:]].value
        pointschosen = worksheet["$W" + cell[2:]].value
        # booster = 1
        # while pointsawarded is None:
        #     pointsawarded = worksheet["$O" + str(int(cell[3:]) - booster)].value
        #     original_practicetext = worksheet["$L" + cell[2:]].value
        #     booster += 1

        errorcheck = [worksheet["$S" + cell[2:]].value, worksheet["$T" + cell[2:]].value]
        if errorcheck[0] == True and errorcheck[1] == False:
            errorcheck = "Points Available - no error"
        elif errorcheck[0] == False and errorcheck[1] == True:
            errorcheck = "Points Unavailable - Error"
        elif errorcheck[0] == True and errorcheck[1] == True:
            errorcheck = "Points Available - Error"
        elif errorcheck[0] == False and errorcheck[1] == False:
            errorcheck = "Points Unavailable - no error"

        try:
            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                        worksheet[ \
                            "$D" + cell[2:]].value + " " + worksheet["$E" + cell[2:]].value + " " + worksheet[ \
                            "$F" + cell[2:]].value
        except TypeError:
            try:  # 4
                practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                            worksheet["$D" + cell[2:]].value + " " + worksheet["$E" + cell[2:]].value
            except TypeError:
                try:  # 3
                    practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                                worksheet["$D" + cell[2:]].value
                except TypeError:
                    try:  # 3
                        practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                            "$E" + cell[2:]].value + " " + \
                                    worksheet["$F" + cell[2:]].value
                    except TypeError:
                        try:  # 3
                            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[ \
                                "$D" + cell[2:]].value + " " + \
                                        worksheet["$E" + cell[2:]].value
                        except TypeError:
                            try:
                                practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                    "$E" + cell[2:]].value
                            except TypeError:
                                try:
                                    practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[ \
                                        "$D" + cell[2:]].value + " " + \
                                                worksheet["$E" + cell[2:]].value
                                except TypeError:
                                    try:
                                        practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                            "$C" + cell[2:]].value
                                    except TypeError:
                                        try:
                                            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                                "$F" + cell[2:]].value
                                        except TypeError:
                                            try:
                                                practices = str(worksheet["$B" + cell[2:]].value)
                                            except:
                                                practices = "No Practice Given"
        if RorFcheck is None:
            value += 1
            continue
        elif RorFcheck is not None and mandatoryitem is True:
            header.append((practices, mandatoryitem, pointschosen, RorFcheck, errorcheck, practicetext))
            row += 1
        value += 1
    return header
        # logging.info("Practice: " + practices)
def practiceextraction(value, row, worksheet):
    header = []
    while value <= 2500:  # cellrow < 2500:
        criteria = ""
        cell = "$W$" + str(value)

        RorFcheck = [worksheet["$P" + cell[2:]].value, worksheet["$Q" + cell[2:]].value]
        if RorFcheck[0] is None or "X" in str(RorFcheck[0]) or RorFcheck[1] is None or "X" in str(RorFcheck[1]):
            RorFcheck = None
        elif RorFcheck[0] == True and RorFcheck[1] == False:
            RorFcheck = "Rough"
        elif RorFcheck[0] == False and RorFcheck[1] == True:
            RorFcheck = "Final"
        elif RorFcheck[0] == True and RorFcheck[1] == True:
            RorFcheck = "Rough/Final"
        elif RorFcheck[0] == False and RorFcheck[1] == False:
            RorFcheck = "Error Error Error"


        pointsawarded = worksheet["$O" + cell[2:]].value,
        practicetext = worksheet["$L" + cell[2:]].value
        pointschosen = worksheet["$W" + cell[2:]].value
        original_practicetext = worksheet["$L" + cell[2:]].value
        booster = 1
        while pointsawarded is None:
            pointsawarded = worksheet["$O" + str(int(cell[3:]) - booster)].value
            original_practicetext = worksheet["$L" + cell[2:]].value
            booster += 1
        # if isinstance(pointsawarded, int) == False:
        #     pointsawarded = 1
        # elif pointsawarded >0:
        #     pointsawarded=1
        # elif pointsawarded ==0:
        #     pointsawarded = 0


        errorcheck = [worksheet["$S" + cell[2:]].value, worksheet["$T" + cell[2:]].value]
        if errorcheck[0] == True and errorcheck[1] == False:
            errorcheck = "Points Available - no error"
        elif errorcheck[0] == False and errorcheck[1] == True:
            errorcheck = "Points Unavailable - Error"
        elif errorcheck[0] == True and errorcheck[1] == True:
            errorcheck = "Points Available - Error"
        elif errorcheck[0] == False and errorcheck[1] == False:
            errorcheck = "Points Unavailable - no error"

        try:
            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                        worksheet[ \
                            "$D" + cell[2:]].value + " " + worksheet["$E" + cell[2:]].value + " " + worksheet[ \
                            "$F" + cell[2:]].value
        except TypeError:
            try:  # 4
                practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                            worksheet["$D" + cell[2:]].value + " " + worksheet["$E" + cell[2:]].value
            except TypeError:
                try:  # 3
                    practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet["$C" + cell[2:]].value + " " + \
                                worksheet["$D" + cell[2:]].value
                except TypeError:
                    try:  # 3
                        practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                            "$E" + cell[2:]].value + " " + \
                                    worksheet["$F" + cell[2:]].value
                    except TypeError:
                        try:  # 3
                            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[ \
                                "$D" + cell[2:]].value + " " + \
                                        worksheet["$E" + cell[2:]].value
                        except TypeError:
                            try:
                                practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                    "$E" + cell[2:]].value
                            except TypeError:
                                try:
                                    practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[ \
                                        "$D" + cell[2:]].value + " " + \
                                                worksheet["$E" + cell[2:]].value
                                except TypeError:
                                    try:
                                        practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                            "$C" + cell[2:]].value
                                    except TypeError:
                                        try:
                                            practices = str(worksheet["$B" + cell[2:]].value) + " " + worksheet[
                                                "$F" + cell[2:]].value
                                        except TypeError:
                                            try:
                                                practices = str(worksheet["$B" + cell[2:]].value)
                                            except:
                                                practices = "No Practice Given"

        # logging.info("Practice: " + practices)
        if RorFcheck is None:
            value += 1
            continue
        elif RorFcheck is not None and pointsawarded != 0 and pointschosen is not None and pointschosen is not False:
            print(pointsawarded)
            header.append((practices, pointsawarded, pointschosen, RorFcheck, errorcheck, practicetext, original_practicetext))
            row += 1
        value += 1
    return header


def dissect(file):
    try:
        workbook = openpyxl.load_workbook((file), data_only=True)
        worksheet = workbook['Verification Report']
    except:
        return
    basefilename = os.path.basename(file)
    if ".xlsm" in basefilename:
        year = "2012"
    elif ".xlsx" in basefilename:
        yearchecksheet = workbook['Info & Intro']
        yearcheck = yearchecksheet['$D$1'].value
        if "2015" in yearcheck:
            year = "2015"
        elif "2020" in yearcheck:
            year = "2020"
    # return year
    if worksheet["X5"].value is None:
        xlname = worksheet["X4"].value
    else:
        xlname = worksheet["X4"].value + " " + str(worksheet["X5"].value)
    try:
        xlname = xlname.replace(":", "")
        xlname = xlname.replace("/", "")
        xlname = xlname.replace("\\", "")
    except:
        return
    value = 14
    row = 0

    pworkbook = openpyxl.Workbook()
    psheet = pworkbook.active

    header = practiceextraction(value,row,worksheet)
    header2 = mandatoryextraction(value,row,worksheet)
    #print(header2)
    for topics, headerset in enumerate(header):
        try:
            logging.info(headerset)
            psheet.append(headerset)
        except:
            continue

    for topics, headerset in enumerate(header2):
        try:
            logging.info(headerset)
            psheet.append(headerset)
        except:
            continue
    print(os.path.dirname(os.path.realpath(file)))
    pworkbook.save(os.path.dirname(os.path.realpath(file))+'/practice examiner/'+xlname+'.xlsx')
    return year






def macroYYYY(original_file, renamed_file, file):

    year = dissect(file)
    if year == "2012":
        print("lemon")
    elif year == "2015":
        code = '''        
                sub marine()
    
                Dim i As Long, hasPhotoReviewWorksheet As Boolean
                Dim xlBook2  As Excel.Workbook
                Dim xlSheet2 As Excel.Worksheet
                Workbooks.Open filename:=''' + "\"" + original_file + "\"" + '''
                set xlbook = ActiveWorkbook
                For i = 1 To xlBook.Worksheets.Count
                    If xlBook.Worksheets(i).Name = "Photo Review" Then
                        Debug.Print xlBook.Worksheets(i).Name
                        hasPhotoReviewWorksheet = True
                        Exit For
                    End If
                Next i
    
                If hasPhotoReviewWorksheet Then
                    'Set xlSheet = xlBook.Worksheets("Photo Review")
                    'continue processing
                Else
                    xlBook.Unprotect ("4gotten!")
                    Set xlBook2 = xlApp.Workbooks.Open("y:\CustomSoftware\Databases\VerifiersCertifications\ExcelTemplates\2015NGBSNewConstructionPhotoReviewTemplate.xlsx")
                    xlBook2.Worksheets("Photo Review").select
                    xlBook2.Worksheets("Photo Review").Copy Before:=xlBook.Sheets("Rough Sig.")  '       xlBook.Sheets("Rough Sig.")
                    xlBook.Protect ("4gotten!")
                    xlBook2.Close False
                End If
    
                For Each ws In xlBook.Worksheets
                    mySheetName = ws.Name
    
                    Debug.Print mySheetName
    
                    Select Case mySheetName
                        Case "Photo Review"
    
                            If hasPhotoReviewWorksheet Then
                                xlBook.Unprotect ("4gotten!")
                                Set xlSheet = xlBook.Worksheets("Photo Review")
                                xlSheet.Visible = xlSheetVisible
                                xlSheet.Unprotect ("4gotten!")
    
                                xlBook.Worksheets("Photo Review").select
                                xlSheet.columns("H:I").EntireColumn.hidden = False
    
                            End If
    
                            xlSheet.Protect ("4gotten!")
                            xlBook.Protect ("4gotten!")
    
    
    
                        Case "Verification Report"
                        If mySheetName = "Verification Report" Then
                            Debug.Print mySheetName
                        End If
    
                        Set xlSheet = xlBook.Worksheets("Verification Report")
                        
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' The TryADifferentPassword on error was addedon 07/27/2018 because  a few workbooks have a bad password
                        '  on the the "Verifification Rpt WorkSheet" The password what applied in upper case when it shouldn't have.
                        '   on error try a different password and resume back at TryADifferentPassword:
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        xlSheet.Unprotect ("4gotten!")
                        xlSheet.Range("A:H").AutoFilter Field:=1
                        xlSheet.Range("A:H").AutoFilter Field:=7
                        xlSheet.Range("A:H").AutoFilter Field:=8
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        TryADifferentPassword:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
                            xlSheet.columns("Z").EntireColumn.hidden = False
                            xlSheet.columns("X:Z").ColumnWidth = 22
    
                            'The reviewers wanted this unlocked so they could easily copy the text and send it
                            ' to a verifier if needed.
                            For r = 9 To 2500
                                If xlSheet.Range("L" & r).MergeCells Then
                                    'Debug.Print r
                                    For Each mrng In xlSheet.Range("L" & r)
                                        With mrng
                                            .MergeArea.Locked = False
                                        End With
                                    Next
                                Else
                                    xlSheet.Range("L" & r).Locked = False
                                End If
                            Next r
    
                            With xlSheet
                                .Protect Password:=("4gotten!"), AllowFiltering:=True
                                .EnableSelection = xlUnlockedCells
                            End With                      
                    End Select
                Next        
    
            ActiveWorkbook.SaveAs filename:=''' + "\"" + renamed_file + "\"" + ''', FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ThisWorkbook.Saved = True
    ActiveWorkbook.close
            end sub
        '''
    elif year == "2020":
        code = '''        
        sub marine()
        Dim i As Long, hasPhotoReviewWorksheet As Boolean
                Dim xlBook2  As Excel.Workbook
                Dim xlSheet2 As Excel.Worksheet
                Workbooks.Open filename:=''' + "\"" + original_file + "\"" + '''
                set xlbook = ActiveWorkbook
                 For i = 1 To xlBook.Worksheets.Count
                    If xlBook.Worksheets(i).Name = "Photo Review" Then
                        Debug.Print xlBook.Worksheets(i).Name
                        hasPhotoReviewWorksheet = True
                        Exit For
                    End If
                Next i
        For Each ws In xlBook.Worksheets
            mySheetName = ws.Name
            
            Debug.Print mySheetName
            
            Select Case mySheetName
                Case "Photo Review"
                    
                    If hasPhotoReviewWorksheet Then
                        xlBook.Unprotect ("4gotten!")
                        Set xlSheet = xlBook.Worksheets("Photo Review")
                        xlSheet.Visible = xlSheetVisible
                        xlSheet.Unprotect ("4gotten!")
                        
                        xlBook.Worksheets("Photo Review").select
                        xlSheet.columns("H:I").EntireColumn.hidden = False

                    End If
                    
                    'xlSheet.Protect ("4gotten!")
                    xlBook.Protect ("4gotten!")
                    
                    
                
                Case "Verification Report"
                
                    If mySheetName = "Verification Report" Then
                        Debug.Print mySheetName
                    End If
                    
                    ' *************************************************************
                    ' Set the worksheet variable to "Verification Report" worksheet
                    ' *************************************************************
                    Set xlSheet = xlBook.Worksheets("Verification Report")
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' The TryADifferentPassword on error was addedon 07/27/2018 because  a few workbooks have a bad password
                    '  on the the "Verifification Rpt WorkSheet" The password what applied in upper case when it shouldn't have.
                    '   on error try a different password and resume back at TryADifferentPassword:
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    xlSheet.Unprotect ("4gotten!")
                    xlSheet.Range("A:H").AutoFilter Field:=1
                    xlSheet.Range("A:H").AutoFilter Field:=7
                    xlSheet.Range("A:H").AutoFilter Field:=8
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TryADifferentPassword:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    xlSheet.columns("Z:AA").EntireColumn.hidden = False
                    xlSheet.columns("X:Z").ColumnWidth = 22
                    xlSheet.columns("AA").ColumnWidth = 6
                    
                    
                    
                    'The reviewers wanted this unlocked so they could easily copy the text and send it
                    '   to a verifier if needed.
                    For r = 9 To 2500
                        If xlSheet.Range("L" & r).MergeCells Then
                            'Debug.Print r
                            For Each mrng In xlSheet.Range("L" & r)
                                With mrng
                                    .MergeArea.Locked = False
                                End With
                            Next
                        Else
                            xlSheet.Range("L" & r).Locked = False
                        End If
                    Next r
                    

                    
                    'The reviewers wanted this unlocked so they could easily copy the text and send it
                    '   to a verifier if needed.
                    For r = 13 To 2500
                        If xlSheet.Range("Z" & r).MergeCells Then
                            For Each mrng In xlSheet.Range("Z" & r)
                                With mrng
                                    .MergeArea.Locked = False
                                End With
                            Next
                        Else
                            xlSheet.Range("Z" & r).Locked = False
                        End If
                    Next r

                    
                    
                    
                    With xlSheet
                        .Protect Password:=("4gotten!"), AllowFiltering:=True
                        .EnableSelection = xlUnlockedCells
                    End With
                    
            End Select
        Next
        

    ActiveWorkbook.SaveAs filename:=''' + "\"" + renamed_file + "\"" + ''', FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ThisWorkbook.Saved = True
    ActiveWorkbook.close
            end sub
'''
    return code


def task():
    #xl.Quit()
    directory_2015 = os.path.dirname(os.path.dirname(os.path.dirname(os.getcwd())))
    print(directory_2015)
    for files in glob.glob(directory_2015 + "/*.xlsx", recursive=False):
        print(files)
        print(files[:-5])
        reviewfiles = glob.glob(files[:-5] + "*_REVIEW.xlsx", recursive=False)
        # check if any other files have the same code extension as other files
        creation = os.path.getmtime(files)
        print(creation)
        print(reviewfiles)
        print("evaluating")
        if "_REVIEW" not in files and not reviewfiles:
            # #if there is no reviewer file with the same code name, then create and name the
            # reviewer version of the file
            print("now running")
            original_file = files
            renamed_file = files[:-5] + "_REVIEW.xlsx"
            ss = xl.Workbooks.Open(os.path.abspath("excelsheet3.xlsm"), ReadOnly=1)
            try:
                # remove the module from the 2nd derivative file
                for i in ss.VBProject.VBComponents:
                    old_xlmodule = ss.VBProject.VBComponents(i.Name)
                    if old_xlmodule.Type in [1, 2, 3]:
                        ss.VBProject.VBComponents.Remove(old_xlmodule)
            except: 
                print("module does not exist")
            # add the module
            print("adding module")
            xlmodule = ss.VBProject.VBComponents.Add(1)

            code = macroYYYY(original_file, renamed_file, files)

            #use the appropriate code as the macro base and run the macro
            xlmodule.CodeModule.AddFromString(code)
            xlmodule.Name = 'translate'
            ss.Application.Run('translate.marine')
            ss.Close(SaveChanges=False)
            #xl.Quit()
            os.remove(files)
            # delete the fileafter exiting excel
        else:
            print(reviewfiles)
    root.after(5000, task)

def endbutton():
    print("end")

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Quit()
xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = False

root = tk.Tk()

root.after(5000, task)
root.mainloop()

