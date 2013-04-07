#if !defined(AFX_ADDRESSADDENTRYDLG_H__9917CC43_7ED4_4E8D_A14B_3665C1C8B340__INCLUDED_)
#define AFX_ADDRESSADDENTRYDLG_H__9917CC43_7ED4_4E8D_A14B_3665C1C8B340__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddressAddEntryDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddressAddEntryDlg dialog

class CAddressAddEntryDlg : public CDialog
{
// Construction
public:
	CAddressAddEntryDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddressAddEntryDlg)
	enum { IDD = IDD_DIALOG_ADD_ENTRY };
	CString	m_address;
	CString	m_email;
	CString	m_name;
	CString	m_phone;
	CString	m_note;
	//}}AFX_DATA
	CString m_title;
	CString m_nameOriginal;
	BOOL m_init;


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressAddEntryDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddressAddEntryDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSetfocusAddress();
	afx_msg void OnSetfocusName();
	afx_msg void OnSetfocusNotes();
	afx_msg void OnSetfocusPhone();
	afx_msg void OnSetfocusEmail();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSADDENTRYDLG_H__9917CC43_7ED4_4E8D_A14B_3665C1C8B340__INCLUDED_)
