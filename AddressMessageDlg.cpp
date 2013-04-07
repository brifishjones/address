// AddressMessageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "address.h"
#include "AddressMessageDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressMessageDlg dialog


CAddressMessageDlg::CAddressMessageDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddressMessageDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddressMessageDlg)
	m_initializationCode = _T("");
	m_temporaryKey = _T("");
	m_staticMessage = _T("");
	//}}AFX_DATA_INIT
}


void CAddressMessageDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddressMessageDlg)
	DDX_Text(pDX, IDC_INITIALIZATION_CODE, m_initializationCode);
	DDX_Text(pDX, IDC_TEMPORARY_KEY, m_temporaryKey);
	DDX_Text(pDX, IDC_STATIC_MESSAGE, m_staticMessage);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddressMessageDlg, CDialog)
	//{{AFX_MSG_MAP(CAddressMessageDlg)
	ON_BN_CLICKED(IDC_STATIC_HAZARD, OnStaticHazard)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressMessageDlg message handlers

void CAddressMessageDlg::OnStaticHazard() 
{
	// TODO: Add your control notification handler code here
	CWnd *pCBox;

	pCBox = (CComboBox *)GetDlgItem(IDC_STATIC_INITIALIZATION_CODE);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)GetDlgItem(IDC_INITIALIZATION_CODE);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)GetDlgItem(IDC_STATIC_TEMPORARY_KEY);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)GetDlgItem(IDC_TEMPORARY_KEY);
	pCBox->ShowWindow(SW_SHOW);

	return;
}
