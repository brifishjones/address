#if !defined(AFX_ADDRESSMESSAGEDLG_H__6AAA3478_4BE4_469E_80CA_468F6E2C044A__INCLUDED_)
#define AFX_ADDRESSMESSAGEDLG_H__6AAA3478_4BE4_469E_80CA_468F6E2C044A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddressMessageDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddressMessageDlg dialog

class CAddressMessageDlg : public CDialog
{
// Construction
public:
	CAddressMessageDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddressMessageDlg)
	enum { IDD = IDD_DIALOG_MESSAGE };
	CString	m_initializationCode;
	CString	m_temporaryKey;
	CString	m_staticMessage;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressMessageDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddressMessageDlg)
	afx_msg void OnStaticHazard();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSMESSAGEDLG_H__6AAA3478_4BE4_469E_80CA_468F6E2C044A__INCLUDED_)
