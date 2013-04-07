#if !defined(AFX_ADDRESSLICENSEDLG_H__A62ADEF2_BB78_489A_BF6E_8C55B9680D93__INCLUDED_)
#define AFX_ADDRESSLICENSEDLG_H__A62ADEF2_BB78_489A_BF6E_8C55B9680D93__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddressLicenseDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddressLicenseDlg dialog

class CAddressLicenseDlg : public CDialog
{
// Construction
public:
	CAddressLicenseDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddressLicenseDlg)
	enum { IDD = IDD_DIALOG_LICENSE };
	CString	m_license;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressLicenseDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddressLicenseDlg)
	afx_msg void OnStaticHazard();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSLICENSEDLG_H__A62ADEF2_BB78_489A_BF6E_8C55B9680D93__INCLUDED_)
