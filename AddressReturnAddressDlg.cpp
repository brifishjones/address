// AddressReturnAddressDlg.cpp : implementation file
//

#include "stdafx.h"
#include "address.h"
#include "AddressReturnAddressDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressReturnAddressDlg dialog


CAddressReturnAddressDlg::CAddressReturnAddressDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddressReturnAddressDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddressReturnAddressDlg)
	m_returnAddress = _T("");
	//}}AFX_DATA_INIT
}


void CAddressReturnAddressDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddressReturnAddressDlg)
	DDX_Text(pDX, IDC_RETURN_ADDRESS, m_returnAddress);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddressReturnAddressDlg, CDialog)
	//{{AFX_MSG_MAP(CAddressReturnAddressDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressReturnAddressDlg message handlers
