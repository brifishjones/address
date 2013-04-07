// address.cpp : Defines the class behaviors for the application.
//
#include "stdafx.h"
#include "address.h"
#include "YoLicense.h"   // static library with licensing routines
//#include "addressLic.h"
//#include "math.h"

#include "MainFrm.h"
#include "addressLicenseDlg.h"
#include "addressMessageDlg.h"
#include "addressDoc.h"
#include "addressView.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


/////////////////////////////////////////////////////////////////////////////
// CAddressApp

BEGIN_MESSAGE_MAP(CAddressApp, CWinApp)
	//{{AFX_MSG_MAP(CAddressApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	ON_COMMAND(ID_FILE_PRINT_SETUP, OnFilePrintSetup)
	//}}AFX_MSG_MAP
	// Standard file based document commands
	ON_COMMAND(ID_FILE_NEW, CWinApp::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, CWinApp::OnFileOpen)
	// Standard print setup command
	ON_COMMAND(ID_FILE_PRINT_SETUP, CWinApp::OnFilePrintSetup)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressApp construction

CAddressApp::CAddressApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CAddressApp object

CAddressApp theApp;

// This identifier was generated to be statistically unique for your app.
// You may change it if you prefer to choose a specific identifier.

// {F3CAEE0F-4D24-4800-A746-9B7080362804}
static const CLSID clsid =
{ 0xf3caee0f, 0x4d24, 0x4800, { 0xa7, 0x46, 0x9b, 0x70, 0x80, 0x36, 0x28, 0x2 } };
//{ 0xf3caee0f, 0x4d24, 0x4800, { 0xa7, 0x46, 0x9b, 0x70, 0x80, 0x36, 0x28, 0x4 } };

#define ADDRESS_VERSION "2004.08.17"  // change version constant when creating updated release

class CYoLicense m_license("address");

/////////////////////////////////////////////////////////////////////////////
// CAddressApp initialization

BOOL CAddressApp::InitInstance()
{
	char message[256];
	char progID[48]; 
	
	sprintf(progID, "F3CAEE0F4D244800A7469B7080362802");  // must be same in addressLic.DLL

	m_init = FALSE;
	//m_license("address");

	if (!AfxSocketInit())
	{
		AfxMessageBox(IDP_SOCKETS_INIT_FAILED);
		return FALSE;
	}

	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	TRACE2("%d %x\n", 2004, 2004);
	TRACE2("%d %x\n", 9999, 9999);
	TRACE2("%d %x\n", 59, 59);
	TRACE2("%d %x\n", 23, 23);
	TRACE2("%d %x\n", 999, 999);
	TRACE2("%04d %04x\n", 0, 0);

	DWORD dwSerialNumber;
    if (!::GetVolumeInformation ("C:\\", NULL, 0,
        &dwSerialNumber, NULL, NULL, NULL, 0))
	{
        dwSerialNumber = 0xFFFFFFFF;
	}
	TRACE2("hard disk serial number = %x %d\n", dwSerialNumber, dwSerialNumber);

	if (!m_license.yoValidateLicense(progID, message))
	{
		if (!m_license.yoValidateTemporaryLicense(message))
		{
			if (!rpMessage(message))
			{
				return FALSE;
			}
		}
	}

	//rpInitializationCode();
	//rpInitializeRegistry('a');

	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	// Change the registry key under which our settings are stored.
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization.
	//SetRegistryKey(_T("Local AppWizard-Generated Applications"));
	SetRegistryKey(_T("bjones"));

	LoadStdProfileSettings();  // Load standard INI file options (including MRU)

	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views.

	CSingleDocTemplate* pDocTemplate;
	pDocTemplate = new CSingleDocTemplate(
		IDR_MAINFRAME,
		RUNTIME_CLASS(CAddressDoc),
		RUNTIME_CLASS(CMainFrame),       // main SDI frame window
		RUNTIME_CLASS(CAddressView));
	AddDocTemplate(pDocTemplate);

	RegisterShellFileTypes(FALSE);
	EnableShellOpen();

	// Connect the COleTemplateServer to the document template.
	//  The COleTemplateServer creates new documents on behalf
	//  of requesting OLE containers by using information
	//  specified in the document template.
	m_server.ConnectTemplate(clsid, pDocTemplate, TRUE);
		// Note: SDI applications register server objects only if /Embedding
		//   or /Automation is present on the command line.

	// Parse command line for standard shell commands, DDE, file open
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	// Check to see if launched as OLE server
	if (cmdInfo.m_bRunEmbedded || cmdInfo.m_bRunAutomated)
	{
		// Register all OLE server (factories) as running.  This enables the
		//  OLE libraries to create objects from other applications.
		COleTemplateServer::RegisterAll();

		// Application was run with /Embedding or /Automation.  Don't show the
		//  main window in this case.
		return TRUE;
	}

	// When a server application is launched stand-alone, it is a good idea
	//  to update the system registry in case it has been damaged.
	m_server.UpdateRegistry(OAT_DISPATCH_OBJECT);
	COleObjectFactory::UpdateRegistryAll();

	// Dispatch commands specified on the command line
	if (!ProcessShellCommand(cmdInfo))
		return FALSE;

	// The one and only window has been initialized, so show and update it.
	m_pMainWnd->ShowWindow(SW_SHOW);
	m_pMainWnd->UpdateWindow();

	m_init = TRUE;

	return TRUE;
}




/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	CString	m_license;
	CString	m_version;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	m_license = _T("");
	m_version = _T("");
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	DDX_Text(pDX, IDC_LICENSE, m_license);
	DDX_Text(pDX, IDC_VERSION, m_version);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

// App command to run the dialog
void CAddressApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	char sLicense[64];

	if (m_license.m_license[0] != '\0')
	{
		sprintf(sLicense, "%s", m_license.m_license);
		//j = 0;
		//for (i = 0; i < 16; i++)
		//{
		//	sLicense[j] = m_license.m_license[i];
		//	if (i % 4 == 3)
		//	{
		//		j++;
		//		sLicense[j] = ' ';
		//	}
		//	j++;
		//}
		//sLicense[j] = '\0';
	}
	else if (m_license.m_temporaryKey[0] != '\0')
	{
		sprintf(sLicense, "expires %s",
			m_license.m_temporaryKey);
	}
	else // should never be called
	{
		sLicense[0] = '\0';
	}

	aboutDlg.m_version = "Version: " + (CString)ADDRESS_VERSION;
	aboutDlg.m_license = "License: " + (CString)sLicense;
	aboutDlg.DoModal();
	return;
}

/////////////////////////////////////////////////////////////////////////////
// CAddressApp message handlers


void CAddressApp::OnFilePrintSetup() 
{
	CWinApp::OnFilePrintSetup();

	CWnd *pWnd = GetMainWnd( );
	CView *pView = ((CMainFrame *)pWnd)->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();

	CDC dc;
	CPrintDialog dlg(FALSE);
	CWnd *pCBox;
	CWinApp* app = AfxGetApp();

	app->GetPrinterDeviceDefaults(&dlg.m_pd);

	LPDEVMODE lp = (LPDEVMODE) ::GlobalLock(dlg.m_pd.hDevMode);
	ASSERT(lp);

	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_paperSizeBook = lp->dmPaperSize;
		pDoc->m_paperOrientationBook = lp->dmOrientation;
	}
	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_paperSizeList = lp->dmPaperSize;
		pDoc->m_paperOrientationList = lp->dmOrientation;
	}
	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		if (lp->dmPaperSize == DMPAPER_LETTER || lp->dmPaperSize == DMPAPER_A4)
		{
			pDoc->m_paperSizeLabel = lp->dmPaperSize;

			if (lp->dmPaperSize == DMPAPER_LETTER)
			{
				pDoc->m_labelLeftMargin = 12;
				pDoc->m_labelTopMargin = 50;
				pDoc->m_labelWidth = 262;
				pDoc->m_labelHeight = 100;
				pDoc->m_labelColumnWidth = 278;
				pDoc->m_labelDimension[0] = 3;
				pDoc->m_labelDimension[1] = 10;
			}
			else if (lp->dmPaperSize == DMPAPER_A4)
			{
				pDoc->m_labelLeftMargin = 13;
				pDoc->m_labelTopMargin = 51;
				pDoc->m_labelWidth = 252;
				pDoc->m_labelHeight = 133;
				pDoc->m_labelColumnWidth = 270;
				pDoc->m_labelDimension[0] = 3;
				pDoc->m_labelDimension[1] = 8;
			}
			//pDoc->m_paperOrientationLabel = lp->dmOrientation;
		}
	}
	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_paperSizeEnvelope = lp->dmPaperSize;
		pDoc->m_paperOrientationEnvelope = lp->dmOrientation;
	}

	::GlobalUnlock(dlg.m_pd.hDevMode);

	dlg.CreatePrinterDC();
	dc.Attach(dlg.m_pd.hDC);


	pDoc->rpComputeDrawPositions();

	return;
}

BOOL CAddressApp::OnIdle(LONG lCount) 
{
	if (m_init == FALSE)
	{
		CWinApp::OnIdle(lCount);
	}

	CWnd *pWnd = GetMainWnd( );
	CView *pView = ((CMainFrame *)pWnd)->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	CWnd *pCBox;
	int i;
	int lineSpacing;
	int columnWidth;
	int returnSize;
	char s[32];
	CPoint pt;

	// if edit box has focus and value has changed, reflect update in CView
	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LINE_SPACING);
		if (((CMainFrame *)pWnd)->GetFocus() == pCBox)
		{
			pCBox->GetWindowText(s, 32);
			i = sscanf(s, "%d", &lineSpacing);
			if (i == 1 && pDoc->m_lineSpacingList != lineSpacing && lineSpacing >= 8)
			{
				lineSpacing = MAX(lineSpacing, 8);
				lineSpacing = MIN(lineSpacing, 72);
				pDoc->m_lineSpacingList = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
				sprintf(s, "%d", pDoc->m_lineSpacingList);
				pCBox->SetWindowText(s);
				pCBox->UpdateWindow();
				((CEdit *)pCBox)->SetSel(0, pCBox->GetWindowTextLength(), FALSE);

			}
		}
	}

	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LINE_SPACING);
		if (((CMainFrame *)pWnd)->GetFocus() == pCBox)
		{
			pCBox->GetWindowText(s, 32);
			i = sscanf(s, "%d", &lineSpacing);
			if (i == 1 && pDoc->m_lineSpacing != lineSpacing && lineSpacing >= 8)
			{
				lineSpacing = MAX(lineSpacing, 8);
				lineSpacing = MIN(lineSpacing, 72);
				pDoc->m_lineSpacing = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
				sprintf(s, "%d", pDoc->m_lineSpacing);
				pCBox->SetWindowText(s);
				pCBox->UpdateWindow();
				((CEdit *)pCBox)->SetSel(0, pCBox->GetWindowTextLength(), FALSE);
			}
		}
		pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_COLUMN_WIDTH);
		if (((CMainFrame *)pWnd)->GetFocus() == pCBox)
		{
			pCBox->GetWindowText(s, 32);
			i = sscanf(s, "%d", &columnWidth);
			if (i == 1 && pDoc->m_columnWidth != columnWidth && columnWidth >= 100)
			{
				columnWidth = MAX(columnWidth, 100);
				columnWidth = MIN(columnWidth,
					pDoc->m_pageWidth - pDoc->m_margin[0] - pDoc->m_margin[2]);
				pDoc->m_columnWidth = columnWidth;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
				sprintf(s, "%d", pDoc->m_columnWidth);
				pCBox->SetWindowText(s);
				pCBox->UpdateWindow();
				((CEdit *)pCBox)->SetSel(0, pCBox->GetWindowTextLength(), FALSE);
			}
		}
	}

	pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_LINE_SPACING);
		if (((CMainFrame *)pWnd)->GetFocus() == pCBox)
		{
			pCBox->GetWindowText(s, 32);
			i = sscanf(s, "%d", &lineSpacing);
			if (i == 1 && pDoc->m_lineSpacingEnvelope != lineSpacing && lineSpacing >= 8)
			{
				lineSpacing = MAX(lineSpacing, 8);
				lineSpacing = MIN(lineSpacing, 72);
				pDoc->m_lineSpacingEnvelope = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
				sprintf(s, "%d", pDoc->m_lineSpacingEnvelope);
				pCBox->SetWindowText(s);
				pCBox->UpdateWindow();
				((CEdit *)pCBox)->SetSel(0, pCBox->GetWindowTextLength(), FALSE);
			}
		}

		pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_RETURN);
		if (((CButton *)pCBox)->GetCheck())
		{
			pCBox = (CComboBox*)(((CMainFrame *)pWnd)->m_wndDlgBarTop).GetDlgItem(IDC_RETURN_SIZE);
			if (((CMainFrame *)pWnd)->GetFocus() == pCBox)
			{
				pCBox->GetWindowText(s, 32);
				i = sscanf(s, "%d", &returnSize);
				if (i == 1 && pDoc->m_returnSizeEnvelope != returnSize && returnSize >= 20)
				{
					returnSize = MAX(returnSize, 20);
					returnSize = MIN(returnSize, 100);
					pDoc->m_returnSizeEnvelope = returnSize;
					pDoc->rpComputeDrawPositions();
					if (pDoc->m_ptrTableList != NULL)
					{
						pDoc->SetModifiedFlag(TRUE);
					}
					sprintf(s, "%d", pDoc->m_returnSizeEnvelope);
					pCBox->SetWindowText(s);
					pCBox->UpdateWindow();
					((CEdit *)pCBox)->SetSel(0, pCBox->GetWindowTextLength(), FALSE);
				}
			}
		}


	}

	return CWinApp::OnIdle(lCount);
}


BOOL CAddressApp::rpMessage(char sMessage[256])
// displays message box containing licensing error
// option exists to enter a temporary key based on initialization code
// if temporary key is valid TRUE is returned else FALSE is returned.
{
	CAddressMessageDlg messageDlg;
	char s[24];
	char *t;
	int i;
	int j;

	// generate initialization code
	m_license.yoGenerateInitializationCode();


	// add spaces to initialization code for easier reading
	j = 0;
	for (i = 0; i < 16; i++)
	{
		s[j] = m_license.m_initializationCode[i];
		j++;
		if (i < 15 && (i % 4) == 3)
		{
			s[j] = ' ';
			j++;
		}
	}
	s[j] = '\0';

	messageDlg.m_staticMessage.Format("%s", (LPCTSTR)
		"Unable to run address application due to licensing error.  A Temporary license may be activated by entering a temporary key based on the given initialization code.");
	messageDlg.m_initializationCode.Format("%s", (LPCTSTR)s);
	messageDlg.DoModal();

	// generate the temporary key
	m_license.yoTemporaryKey();

	sprintf(s, "%s", messageDlg.m_temporaryKey);
	// remove spaces in temporary key
	i = 0;
	for (t = s; i < 24 && *t != '\0'; t++)
	{
		for (; *t == ' '; t++);
		s[i] = *t;
		i++;
	}
	s[16] = '\0';
	
	// find placeholder for expiration code
	i = (((int)m_license.m_temporaryKey[0] - 97) % 14) + 1;
	// given the placeholder, insert expiration code
	// expiration codes
	// a = 360 days
	// b = 180 days
	// c = 90 days
	// d = 60 days
	// e = 30 days
	// f = 7 days
	// g and other chars = 1 day
	m_license.m_temporaryKey[i] = s[i];

	// if temporary key entered in dialog box matches generated temporary key
	// then initialize temporary licensing
	if (strcmp(s, m_license.m_temporaryKey) == 0)
	{
		return m_license.yoInitializeTemporaryLicense(s[i], sMessage);
	}
	else
	{
		m_license.m_temporaryKey[0] = '\0';
		return FALSE;
	}
}
