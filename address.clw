; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CAddressAddEntryDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "address.h"
ODLFile=address.odl
LastPage=0

ClassCount=10
Class1=CAddressApp
Class2=CAddressDoc
Class3=CAddressView
Class4=CMainFrame

ResourceCount=12
Resource1=IDD_ABOUTBOX
Resource2="IDD_PERFORM_LEFT"
Class5=CAboutDlg
Resource3=IDD_DIALOG_LICENSE
Resource4=IDD_DIALOG_RETURN_ADDRESS
Class6=CAddressAddEntryDlg
Resource5=IDR_MAINFRAME
Class7=CAddressReturnAddressDlg
Resource6=IDD_DIALOG_GROUP
Class8=CAddressGroupDlg
Resource7=IDD_DIALOG_ADD_ENTRY
Resource8="IDD_PERFORM_TOP"
Resource9=IDD_ADDRESS_LEFT
Class9=CAddressLicenseDlg
Resource10=IDD_DIALOG_MESSAGE
Class10=CAddressMessageDlg
Resource11=IDD_ADDRESS_TOP
Resource12=IDD_DIALOG_LICENSE_MESSAGE

[CLS:CAddressApp]
Type=0
HeaderFile=address.h
ImplementationFile=address.cpp
Filter=N
LastObject=CAddressApp
BaseClass=CWinApp
VirtualFilter=AC

[CLS:CAddressDoc]
Type=0
HeaderFile=addressDoc.h
ImplementationFile=addressDoc.cpp
Filter=N
LastObject=ID_EDIT_PASTE
BaseClass=CDocument
VirtualFilter=DC

[CLS:CAddressView]
Type=0
HeaderFile=addressView.h
ImplementationFile=addressView.cpp
Filter=C
BaseClass=CScrollView
VirtualFilter=VWC
LastObject=ID_EDIT_COPY


[CLS:CMainFrame]
Type=0
HeaderFile=MainFrm.h
ImplementationFile=MainFrm.cpp
Filter=T
BaseClass=CFrameWnd
VirtualFilter=fWC
LastObject=CMainFrame




[CLS:CAboutDlg]
Type=0
HeaderFile=address.cpp
ImplementationFile=address.cpp
Filter=D
LastObject=CAboutDlg
BaseClass=CDialog
VirtualFilter=dWC

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=6
Control1=IDC_STATIC,static,1342177283
Control2=IDC_VERSION,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Control5=IDC_STATIC,static,1342308352
Control6=IDC_LICENSE,static,1342308352

[MNU:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_FILE_NEW
Command2=ID_FILE_OPEN
Command3=ID_FILE_MERGE
Command4=ID_FILE_SAVE
Command5=ID_FILE_SAVE_AS
Command6=ID_FILE_WRITE_SELECTED_TO_FILE
Command7=ID_FILE_EXPORT_CSV
Command8=ID_FILE_EXPORT_PALM_CSV
Command9=ID_FILE_EXPORT_MAILRC
Command10=ID_FILE_PRINT
Command11=ID_FILE_PRINT_PREVIEW
Command12=ID_FILE_PRINT_SETUP
Command13=ID_FILE_MRU_FILE1
Command14=ID_APP_EXIT
Command15=ID_EDIT_UNDO
Command16=ID_EDIT_CUT
Command17=ID_EDIT_COPY
Command18=ID_EDIT_PASTE
Command19=ID_VIEW_STATUS_BAR
Command20=ID_VIEW_LAST_NAME_FIRST
Command21=ID_VIEW_RETURN_ADDRESS
Command22=ID_FONT_CHOOSE
Command23=ID_FONT_RESTORE_DEFAULT
Command24=ID_GROUP_UPDATE
Command25=ID_GROUP_ONE
Command26=ID_GROUP_TWO
Command27=ID_GROUP_THREE
Command28=ID_GROUP_FOUR
Command29=ID_GROUP_FIVE
Command30=ID_GROUP_SIX
Command31=ID_GROUP_SEVEN
Command32=ID_GROUP_EIGHT
Command33=ID_GROUP_NINE
Command34=ID_GROUP_TEN
Command35=ID_GROUP_ELEVEN
Command36=ID_GROUP_TWELVE
Command37=ID_APP_ABOUT
CommandCount=37

[ACL:IDR_MAINFRAME]
Type=1
Class=CMainFrame
Command1=ID_FILE_NEW
Command2=ID_FILE_OPEN
Command3=ID_FILE_SAVE
Command4=ID_FILE_PRINT
Command5=ID_EDIT_UNDO
Command6=ID_EDIT_CUT
Command7=ID_EDIT_COPY
Command8=ID_EDIT_PASTE
Command9=ID_EDIT_UNDO
Command10=ID_EDIT_CUT
Command11=ID_EDIT_COPY
Command12=ID_EDIT_PASTE
Command13=ID_NEXT_PANE
Command14=ID_PREV_PANE
CommandCount=14

[DLG:IDD_ADDRESS_TOP]
Type=1
Class=?
ControlCount=12
Control1=IDC_LIST,button,1342275587
Control2=IDC_BOOK,button,1342275587
Control3=IDC_ENVELOPE,button,1342275587
Control4=IDC_LABEL,button,1342275587
Control5=IDC_LINE_SPACING,edit,1350631552
Control6=IDC_COLUMN_WIDTH,edit,1350631552
Control7=IDC_LINE_SPACING_DEFAULTS,button,1342275584
Control8=IDC_COLUMN_WIDTH_DEFAULTS,button,1342275584
Control9=IDC_STATIC,static,1342308352
Control10=IDC_RETURN,button,1073741833
Control11=IDC_RETURN_SIZE,edit,1073807488
Control12=IDC_RETURN_PERCENT,static,1073872896

[DLG:IDD_ADDRESS_LEFT]
Type=1
Class=?
ControlCount=5
Control1=IDC_NAME,listbox,1353777417
Control2=IDC_SELECT_ALL,button,1342177289
Control3=IDC_SELECT_NONE,button,1342177289
Control4=IDC_ADD,button,1342279683
Control5=IDC_DELETE,button,1342279683

[DLG:IDD_DIALOG_ADD_ENTRY]
Type=1
Class=CAddressAddEntryDlg
ControlCount=14
Control1=IDC_NAME,edit,1350631552
Control2=IDC_ADDRESS,edit,1350635716
Control3=IDC_EMAIL,edit,1350635716
Control4=IDC_PHONE,edit,1350635716
Control5=IDC_NOTES,edit,1350635716
Control6=IDOK,button,1342242816
Control7=IDCANCEL,button,1342242816
Control8=IDC_STATIC,static,1342308352
Control9=IDC_STATIC,static,1342308352
Control10=IDC_STATIC,static,1342308352
Control11=IDC_STATIC,static,1342308352
Control12=IDC_STATIC,static,1342308352
Control13=IDC_MESSAGE,static,1342308352
Control14=IDC_STATIC_HAZARD,static,1073741838

[CLS:CAddressAddEntryDlg]
Type=0
HeaderFile=AddressAddEntryDlg.h
ImplementationFile=AddressAddEntryDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=IDOK
VirtualFilter=dWC

[DLG:IDD_DIALOG_RETURN_ADDRESS]
Type=1
Class=CAddressReturnAddressDlg
ControlCount=3
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_RETURN_ADDRESS,edit,1350635716

[CLS:CAddressReturnAddressDlg]
Type=0
HeaderFile=AddressReturnAddressDlg.h
ImplementationFile=AddressReturnAddressDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=CAddressReturnAddressDlg
VirtualFilter=dWC

[DLG:IDD_DIALOG_GROUP]
Type=1
Class=CAddressGroupDlg
ControlCount=12
Control1=IDC_NEW_GROUP,edit,1350631552
Control2=IDC_NEW,button,1476493312
Control3=IDC_GROUP_NAME,listbox,1352728833
Control4=IDC_DELETE,button,1476493312
Control5=IDC_MEMBER_NAME,listbox,1353777417
Control6=IDC_ALL,button,1342177289
Control7=IDC_NONE,button,1342177289
Control8=IDOK,button,1342242817
Control9=IDCANCEL,button,1342242816
Control10=IDC_GROUP_TEXT,static,1342308352
Control11=IDC_STATIC,static,1342308352
Control12=IDC_GROUP_MESSAGE,static,1342308352

[CLS:CAddressGroupDlg]
Type=0
HeaderFile=AddressGroupDlg.h
ImplementationFile=AddressGroupDlg.cpp
BaseClass=CDialog
Filter=D
LastObject=CAddressGroupDlg
VirtualFilter=dWC

[DLG:"IDD_PERFORM_LEFT"]
Type=1
Class=?
ControlCount=5
Control1=IDC_NAME,listbox,1353777417
Control2=IDC_SELECT_ALL,button,1342177289
Control3=IDC_SELECT_NONE,button,1342177289
Control4=IDC_ADD,button,1342279683
Control5=IDC_DELETE,button,1342279683

[DLG:"IDD_PERFORM_TOP"]
Type=1
Class=?
ControlCount=10
Control1=IDC_LIST,button,1342275587
Control2=IDC_STATIC,static,1342308352
Control3=IDC_BOOK,button,1342275587
Control4=IDC_ENVELOPE,button,1342275587
Control5=IDC_LABEL,button,1342275587
Control6=IDC_LINE_SPACING,edit,1350631552
Control7=IDC_COLUMN_WIDTH,edit,1350631552
Control8=IDC_LINE_SPACING_DEFAULTS,button,1342275584
Control9=IDC_COLUMN_WIDTH_DEFAULTS,button,1342275584
Control10=IDC_RETURN,button,1073741833

[DLG:IDD_DIALOG_LICENSE]
Type=1
Class=CAddressLicenseDlg
ControlCount=10
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1073807360
Control3=IDC_STATIC,static,1342308352
Control4=IDC_LICENSE,edit,1342244864
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC_HAZARD,static,1342177550
Control7=IDC_STATIC_INITIALIZATION_CODE,static,1073872896
Control8=IDC_INITIALIZATION_CODE,edit,1073809408
Control9=IDC_TEMPORARY_KEY,edit,1073807360
Control10=IDC_STATIC_TEMPORARY_KEY,static,1073872896

[CLS:CAddressLicenseDlg]
Type=0
HeaderFile=AddressLicenseDlg.h
ImplementationFile=AddressLicenseDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CAddressLicenseDlg

[DLG:IDD_DIALOG_MESSAGE]
Type=1
Class=CAddressMessageDlg
ControlCount=8
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1073807360
Control3=IDC_STATIC_MESSAGE,static,1342308353
Control4=IDC_TEMPORARY_KEY,edit,1342242816
Control5=IDC_STATIC_TEMPORARY_KEY,static,1342308352
Control6=IDC_STATIC_HAZARD,static,1342177550
Control7=IDC_STATIC_INITIALIZATION_CODE,static,1342308352
Control8=IDC_INITIALIZATION_CODE,edit,1342244864

[CLS:CAddressMessageDlg]
Type=0
HeaderFile=AddressMessageDlg.h
ImplementationFile=AddressMessageDlg.cpp
BaseClass=CDialog
Filter=D
VirtualFilter=dWC
LastObject=CAddressMessageDlg

[DLG:IDD_DIALOG_LICENSE_MESSAGE]
Type=1
Class=?
ControlCount=12
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1073807360
Control3=IDC_STATIC,static,1342308352
Control4=IDC_LICENSE,edit,1342244864
Control5=IDC_STATIC,static,1342308352
Control6=IDC_STATIC_HAZARD,static,1342177550
Control7=IDC_STATIC_INITIALIZATION_CODE,static,1342308352
Control8=IDC_INITIALIZATION_CODE,edit,1342244864
Control9=IDC_TEMPORARY_KEY,edit,1342242816
Control10=IDC_STATIC_TEMPORARY_KEY,static,1342308352
Control11=IDC_STATIC,button,1342177287
Control12=IDC_STATIC,button,1342177287

