// AddressAddEntryDlg.cpp : implementation file
//

#include "stdafx.h"
#include "address.h"
#include "AddressAddEntryDlg.h"

#include "addressDoc.h"
#include "addressView.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressAddEntryDlg dialog


CAddressAddEntryDlg::CAddressAddEntryDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddressAddEntryDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddressAddEntryDlg)
	m_address = _T("");
	m_email = _T("");
	m_name = _T("");
	m_phone = _T("");
	m_note = _T("");
	//}}AFX_DATA_INIT
}


void CAddressAddEntryDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddressAddEntryDlg)
	DDX_Text(pDX, IDC_ADDRESS, m_address);
	DDX_Text(pDX, IDC_EMAIL, m_email);
	DDX_Text(pDX, IDC_NAME, m_name);
	DDX_Text(pDX, IDC_PHONE, m_phone);
	DDX_Text(pDX, IDC_NOTES, m_note);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddressAddEntryDlg, CDialog)
	//{{AFX_MSG_MAP(CAddressAddEntryDlg)
	ON_EN_SETFOCUS(IDC_ADDRESS, OnSetfocusAddress)
	ON_EN_SETFOCUS(IDC_NAME, OnSetfocusName)
	ON_EN_SETFOCUS(IDC_NOTES, OnSetfocusNotes)
	ON_EN_SETFOCUS(IDC_PHONE, OnSetfocusPhone)
	ON_EN_SETFOCUS(IDC_EMAIL, OnSetfocusEmail)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressAddEntryDlg message handlers

BOOL CAddressAddEntryDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	((CWnd *)this)->SetWindowText(m_title);
	m_init = FALSE;

	m_nameOriginal = m_name;

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CAddressAddEntryDlg::OnSetfocusAddress() 
{
	// TODO: Add your control notification handler code here

	CWnd *pCBox;
	char s[256];

	pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);
	sprintf(s, "Street, city, state/province/region, postal code, etc");
	pCBox->SetWindowText(s);

	return;
	
}

void CAddressAddEntryDlg::OnSetfocusName() 
{
	// TODO: Add your control notification handler code here

	CWnd *pCBox;
	char s[256];

	pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);	
	if (m_title == "Add entry" && m_init == FALSE)
	{
		sprintf(s, "Enter first name(s) followed by last name");
		m_init = TRUE;
	}
	else
	{
		sprintf(s, "First name(s) followed by last name");
	}
	pCBox->SetWindowText(s);

	pCBox = (CComboBox*)GetDlgItem(IDC_STATIC_HAZARD);
	pCBox->ShowWindow(SW_HIDE);
	
}

void CAddressAddEntryDlg::OnSetfocusNotes() 
{
	// TODO: Add your control notification handler code here

	CWnd *pCBox;
	char s[256];

	pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);
	sprintf(s, "Miscellaneous information");
	pCBox->SetWindowText(s);

	return;
		
}

void CAddressAddEntryDlg::OnSetfocusPhone() 
{
	// TODO: Add your control notification handler code here

	CWnd *pCBox;
	char s[256];

	pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);
	sprintf(s, "Home, work, mobile, fax, etc");
	pCBox->SetWindowText(s);

	return;
		
}

void CAddressAddEntryDlg::OnSetfocusEmail() 
{
	// TODO: Add your control notification handler code here

	CWnd *pCBox;
	char s[256];

	pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);
	sprintf(s, "Email address followed by optional nickname");
	pCBox->SetWindowText(s);

	return;
		
}

void CAddressAddEntryDlg::OnOK() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[256];
	CAddressTable *ptrTable;
	char *t;
	char entry[256];
	int i;

	UpdateData(TRUE);
	//m_name.TrimLeft();
	//m_name.TrimRight();

	//if (m_name.Right(1) == ";")
	//{
	//	m_name.Delete(m_name.GetLength() - 1, 1);
	//	m_name.TrimRight();
	//}

	// massage name variable
	sprintf(s, "%s", (LPCTSTR)m_name);
	i = 0;
	for (t = s; *t == ' '; t++);  // remove leading white space if any
	for (; *t != '\0'; t++)
	{
		if (*t == 0x0D && *(t + 1) == 0x0A)
		{
			entry[i] = ' ';
			t++;
		}
		else
		{
			entry[i] = *t;
		}
		i++;
	}
	entry[i] = '\0';

	// remove any trailing whitespace or semi-colins
	for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
	{
		*t = '\0';
	}

	m_name = entry;

	// Check to make certain that entered entry name is unique
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (m_name == (CString)ptrTable->m_name && m_nameOriginal != (CString)ptrTable->m_name)
		{
			pCBox = (CComboBox*)GetDlgItem(IDC_MESSAGE);
			sprintf(s, "Name is not unique!  Re-enter name.");
			pCBox->SetWindowText(s);
			pCBox = (CComboBox*)GetDlgItem(IDC_STATIC_HAZARD);
			pCBox->ShowWindow(SW_SHOW);

			return;
		}
	}

	// TODO: Add extra validation here
	TRACE0("CAddressAddEntryDlg::OnOK()\n");
	
	CDialog::OnOK();
}
