#if !defined(AFX_ADDRESSRETURNADDRESSDLG_H__B0FD2760_5CB4_4955_A317_A14D73BED830__INCLUDED_)
#define AFX_ADDRESSRETURNADDRESSDLG_H__B0FD2760_5CB4_4955_A317_A14D73BED830__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddressReturnAddressDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddressReturnAddressDlg dialog

class CAddressReturnAddressDlg : public CDialog
{
// Construction
public:
	CAddressReturnAddressDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddressReturnAddressDlg)
	enum { IDD = IDD_DIALOG_RETURN_ADDRESS };
	CString	m_returnAddress;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressReturnAddressDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddressReturnAddressDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSRETURNADDRESSDLG_H__B0FD2760_5CB4_4955_A317_A14D73BED830__INCLUDED_)
