

   MEMBER('ExcelMergeExample.clw')                         ! This is a MEMBER module


   INCLUDE('ABTOOLBA.INC'),ONCE
   INCLUDE('ABUTIL.INC'),ONCE
   INCLUDE('ABWINDOW.INC'),ONCE

                     MAP
                       INCLUDE('EXCELMERGEEXAMPLE001.INC'),ONCE        !Local module procedure declarations
                     END


!!! <summary>
!!! Generated from procedure template - Window
!!! </summary>
Main PROCEDURE 

TemplateFile         CSTRING(255)                          ! 
GeneratedFile        CSTRING(255)                          ! 
Window               WINDOW('Excel Merge Example'),AT(,,433,87),FONT('Segoe UI',9),AUTO,ICON(ICON:Clarion),GRAY, |
  SYSTEM,IMM
                       BUTTON('Close'),AT(393,69),USE(?Close)
                       PROMPT('Template File:'),AT(8,10),USE(?TemplateFile:Prompt)
                       ENTRY(@s255),AT(61,9,348,10),USE(TemplateFile)
                       PROMPT('Generated File:'),AT(8,24),USE(?GeneratedFile:Prompt)
                       ENTRY(@s255),AT(61,23,348,10),USE(GeneratedFile)
                       BUTTON('...'),AT(413,8,12,12),USE(?LookupFile)
                       BUTTON('...'),AT(413,22,12,12),USE(?LookupFile:2)
                       BUTTON('Generate'),AT(8,43),USE(?GenerateBtn)
                     END

ThisWindow           CLASS(WindowManager)
Init                   PROCEDURE(),BYTE,PROC,DERIVED
Kill                   PROCEDURE(),BYTE,PROC,DERIVED
TakeAccepted           PROCEDURE(),BYTE,PROC,DERIVED
                     END

Toolbar              ToolbarClass
FileLookup2          SelectFileClass
FileLookup3          SelectFileClass

  CODE
  GlobalResponse = ThisWindow.Run()                        ! Opens the window and starts an Accept Loop

!---------------------------------------------------------------------------
DefineListboxStyle ROUTINE
!|
!| This routine create all the styles to be shared in this window
!| It`s called after the window open
!|
!---------------------------------------------------------------------------

ThisWindow.Init PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  GlobalErrors.SetProcedureName('Main')
  SELF.Request = GlobalRequest                             ! Store the incoming request
  ReturnValue = PARENT.Init()
  IF ReturnValue THEN RETURN ReturnValue.
  SELF.FirstField = ?Close
  SELF.VCRRequest &= VCRRequest
  SELF.Errors &= GlobalErrors                              ! Set this windows ErrorManager to the global ErrorManager
  ! Restore preserved local variables from non-volatile store
  TemplateFile = INIMgr.TryFetch('Main_PreservedVars','TemplateFile')
  GeneratedFile = INIMgr.TryFetch('Main_PreservedVars','GeneratedFile')
  SELF.AddItem(Toolbar)
  CLEAR(GlobalRequest)                                     ! Clear GlobalRequest after storing locally
  CLEAR(GlobalResponse)
  IF SELF.Request = SelectRecord
     SELF.AddItem(?Close,RequestCancelled)                 ! Add the close control to the window manger
  ELSE
     SELF.AddItem(?Close,RequestCompleted)                 ! Add the close control to the window manger
  END
  SELF.Open(Window)                                        ! Open window
  Do DefineListboxStyle
  INIMgr.Fetch('Main',Window)                              ! Restore window settings from non-volatile store
  FileLookup2.Init
  FileLookup2.ClearOnCancel = True
  FileLookup2.Flags=BOR(FileLookup2.Flags,FILE:LongName)   ! Allow long filenames
  FileLookup2.SetMask('All Files','*.*')                   ! Set the file mask
  FileLookup3.Init
  FileLookup3.ClearOnCancel = True
  FileLookup3.Flags=BOR(FileLookup3.Flags,FILE:LongName)   ! Allow long filenames
  FileLookup3.Flags=BOR(FileLookup3.Flags,FILE:Directory)  ! Allow Directory Dialog
  FileLookup3.SetMask('All Files','*.*')                   ! Set the file mask
  SELF.SetAlerts()
  RETURN ReturnValue


ThisWindow.Kill PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  ReturnValue = PARENT.Kill()
  IF ReturnValue THEN RETURN ReturnValue.
  IF SELF.Opened
    INIMgr.Update('Main',Window)                           ! Save window data to non-volatile store
  END
  ! Save preserved local variables in non-volatile store
  INIMgr.Update('Main_PreservedVars','TemplateFile',TemplateFile)
  INIMgr.Update('Main_PreservedVars','GeneratedFile',GeneratedFile)
  GlobalErrors.SetProcedureName
  RETURN ReturnValue


ThisWindow.TakeAccepted PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receive all EVENT:Accepted's
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeAccepted()
    CASE ACCEPTED()
    OF ?LookupFile
      ThisWindow.Update()
      TemplateFile = FileLookup2.Ask(1)
      DISPLAY
    OF ?LookupFile:2
      ThisWindow.Update()
      GeneratedFile = FileLookup3.Ask(1)
      DISPLAY
    OF ?GenerateBtn
      ThisWindow.Update()
      CreateGeneratedFile(TemplateFile,GeneratedFile)  
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue

