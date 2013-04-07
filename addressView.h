// addressView.h : interface of the CAddressView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_ADDRESSVIEW_H__A72117EA_C87D_48FE_A118_71667BCD614B__INCLUDED_)
#define AFX_ADDRESSVIEW_H__A72117EA_C87D_48FE_A118_71667BCD614B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CAddressView : public CScrollView
{
protected: // create from serialization only
	CAddressView();
	DECLARE_DYNCREATE(CAddressView)

// Attributes
public:
	CAddressDoc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressView)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual void OnInitialUpdate();
	virtual void OnPrepareDC(CDC* pDC, CPrintInfo* pInfo = NULL);
	protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnUpdate(CView* pSender, LPARAM lHint, CObject* pHint);
	virtual void OnPrint(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrintPreview(CDC* pDC, CPrintInfo* pInfo, POINT point, CPreviewView* pView);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CAddressView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

public:
	//int CAddressView::rpComputeDrawPositions();
	BOOL rpLeftDlgListbox(void);
	void rpDrawPageHeader(CDC *, CPrintInfo *);

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CAddressView)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnFilePrintPreview();
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in addressView.cpp
inline CAddressDoc* CAddressView::GetDocument()
   { return (CAddressDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSVIEW_H__A72117EA_C87D_48FE_A118_71667BCD614B__INCLUDED_)
