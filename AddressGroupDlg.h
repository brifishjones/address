#if !defined(AFX_ADDRESSGROUPDLG_H__08A45CA2_7166_411E_B910_188D47A9E27C__INCLUDED_)
#define AFX_ADDRESSGROUPDLG_H__08A45CA2_7166_411E_B910_188D47A9E27C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddressGroupDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddressGroupDlg dialog

class CAddressGroupDlg : public CDialog
{
// Construction
public:
	CAddressGroupDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddressGroupDlg)
	enum { IDD = IDD_DIALOG_GROUP };
	CListBox	m_memberName;
	CListBox	m_groupName;
	//}}AFX_DATA

	class CAddressGroup *m_ptrGroupListDup;
	BOOL m_init; // needed to skip OnSetfocusNewGroup() when dialog is initialized.


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressGroupDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddressGroupDlg)
	afx_msg void OnUpdateNewGroup();
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeGroupName();
	afx_msg void OnSelchangeMemberName();
	afx_msg void OnNew();
	afx_msg void OnAll();
	afx_msg void OnNone();
	afx_msg void OnDelete();
	afx_msg void OnFileExportCsv();
	afx_msg void OnKillfocusMemberName();
	afx_msg void OnSetfocusMemberName();
	afx_msg void OnKillfocusGroupName();
	afx_msg void OnSetfocusGroupName();
	afx_msg void OnSetfocusNewGroup();
	afx_msg void OnKillfocusNewGroup();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSGROUPDLG_H__08A45CA2_7166_411E_B910_188D47A9E27C__INCLUDED_)
