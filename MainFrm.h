// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAINFRM_H__B506EDEC_9BF1_4AF0_BE81_62791EED96F3__INCLUDED_)
#define AFX_MAINFRM_H__B506EDEC_9BF1_4AF0_BE81_62791EED96F3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CAddressProgram;
class CAddressTable;
class CAddressAddEntryDlg;

class CMainFrame : public CFrameWnd
{
	
protected: // create from serialization only
	CMainFrame();
	DECLARE_DYNCREATE(CMainFrame)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMainFrame)
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CStatusBar m_wndStatusBar;
public:
	CDialogBar m_wndDlgBarTop;
	CDialogBar m_wndDlgBarLeft;
	CAddressProgram *m_curProgram;
	BOOL m_allGroupMembersSelected[30];

// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnViewReturnAddress();
	afx_msg void OnFontChoose();
	afx_msg void OnFontRestoreDefault();
	afx_msg void OnGroupUpdate();
	afx_msg void OnUpdateGroupUpdate(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupOne(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupFive(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupFour(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupSeven(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupSix(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupThree(CCmdUI* pCmdUI);
	afx_msg void OnUpdateGroupTwo(CCmdUI* pCmdUI);
	afx_msg void OnGroupFive();
	afx_msg void OnGroupFour();
	afx_msg void OnGroupOne();
	afx_msg void OnGroupSeven();
	afx_msg void OnGroupSix();
	afx_msg void OnGroupThree();
	afx_msg void OnGroupTwo();
	afx_msg void OnViewLastNameFirst();
	afx_msg void OnUpdateViewLastNameFirst(CCmdUI* pCmdUI);
	afx_msg void OnGroupEight();
	afx_msg void OnUpdateGroupEight(CCmdUI* pCmdUI);
	afx_msg void OnGroupNine();
	afx_msg void OnUpdateGroupNine(CCmdUI* pCmdUI);
	afx_msg void OnGroupTen();
	afx_msg void OnUpdateGroupTen(CCmdUI* pCmdUI);
	afx_msg void OnGroupEleven();
	afx_msg void OnUpdateGroupEleven(CCmdUI* pCmdUI);
	afx_msg void OnGroupTwelve();
	afx_msg void OnUpdateGroupTwelve(CCmdUI* pCmdUI);
	afx_msg void OnEditCopy();
	afx_msg void OnUpdateEditCopy(CCmdUI* pCmdUI);
	//}}AFX_MSG
public:
	afx_msg void OnList(void);
	afx_msg void OnBook(void);
	afx_msg void OnEnvelope(void);
	afx_msg void OnLabel(void);
	afx_msg void OnReturn(void);
	afx_msg void OnLineSpacingDefaults(void);
	afx_msg void OnColumnWidthDefaults(void);
	afx_msg void OnSelectAll(void);
	afx_msg void OnSelectNone(void);
	afx_msg void OnAddEntry(void);
	afx_msg void OnDelete(void);
	afx_msg void OnName(void);
	afx_msg void OnNameEditEntry(void);
	afx_msg void OnUpdateListUI(CCmdUI *);
	afx_msg void OnUpdateBookUI(CCmdUI *);
	afx_msg void OnUpdateEnvelopeUI(CCmdUI *);
	afx_msg void OnUpdateLabelUI(CCmdUI *);
	afx_msg void OnLineSpacing(void);
	afx_msg void OnColumnWidth(void);


	DECLARE_MESSAGE_MAP()

// helper functions
public:
	void rpReadEntryDlg(CAddressTable *, CAddressAddEntryDlg *);
	void rpWriteEntryDlg(CAddressTable *, CAddressAddEntryDlg *);
	BOOL rpUpdateGroupMenu(void);
	BOOL rpSelectGroupMembers(int);
	void rpCopyToClipboard(char *s);
};





/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__B506EDEC_9BF1_4AF0_BE81_62791EED96F3__INCLUDED_)
