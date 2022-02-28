import os
import win32com.client as win32
import tkinter as tk
import glob
import openpyxl

def dissect(file):
    workbook = openpyxl.load_workbook((file), data_only=True)
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
            end sub
'''
    return code


def task():
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
xl.Visible = False

root = tk.Tk()

root.after(5000, task)
root.mainloop()

