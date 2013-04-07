// AddressLicenseDlg.cpp : implementation file
//

#include "stdafx.h"
#include "address.h"
#include "AddressLicenseDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressLicenseDlg dialog


CAddressLicenseDlg::CAddressLicenseDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddressLicenseDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddressLicenseDlg)
	m_license = _T("");
	//}}AFX_DATA_INIT
}


void CAddressLicenseDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddressLicenseDlg)
	DDX_Text(pDX, IDC_LICENSE, m_license);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddressLicenseDlg, CDialog)
	//{{AFX_MSG_MAP(CAddressLicenseDlg)
	ON_BN_CLICKED(IDC_STATIC_HAZARD, OnStaticHazard)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressLicenseDlg message handlers

void CAddressLicenseDlg::OnStaticHazard() 
{
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
