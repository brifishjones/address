// addressDoc.h : interface of the CAddressDoc class

#define MIN(x, y)  (((x) < (y)) ? (x) : (y))
#define MAX(x, y)  (((x) > (y)) ? (x) : (y))
#define MAXIMUM_NUMBER_GROUPS 12

//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_ADDRESSDOC_H__41A4E9BA_D2D4_4DA9_B764_21F314A2A46E__INCLUDED_)
#define AFX_ADDRESSDOC_H__41A4E9BA_D2D4_4DA9_B764_21F314A2A46E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CAddressTable;
class CAddressAlias;
class CAddressGroup;

class CAddressDoc : public CDocument
{
protected: // create from serialization only
	CAddressDoc();
	DECLARE_DYNCREATE(CAddressDoc)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	virtual BOOL OnOpenDocument(LPCTSTR lpszPathName);
	virtual BOOL OnSaveDocument(LPCTSTR lpszPathName);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CAddressDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

public:
	void rpFreeDoc(void);
	BOOL rpReadDoc(FILE *);
	BOOL rpReadCsvFile(FILE *);
	BOOL rpReadPalmCsvFile(FILE *);
	void rpRemoveQuotes(char s[256]);
	BOOL rpWriteDoc(LPCTSTR);
	CSize rpGetDocSize(void);
	BOOL rpUpdateAddressList(void);
	int rpComputeDrawPositions(void);
	int rpSetFirstLabel(CPoint);
	void rpSortName(void);
	//void rpSortName(CAddressTable *);
	BOOL rpUpdateGroupMenu(void);
	CAddressTable *m_ptrTableList;
	CAddressAlias *m_ptrAliasList;
	CAddressGroup *m_ptrGroupList;
	CAddressGroup *m_groupSelected;
	char m_returnAddress[256];
	int m_cellSpacing[2];
	int m_tableSize[2];
	int m_offsetBorder;  // not used
	int m_pageWidth;
	int m_pageHeight;
	int m_printableWidth;
	int m_printableHeight;
	int m_physicalOffsetX;  // unprintable area at left of page
	int m_physicalOffsetY;  // unprintable area at top of page
	int m_lineSpacing;
	int m_lineSpacingEnvelope;
	int m_returnSizeEnvelope;
	int m_lineSpacingList;
	CFont m_fontBook;
	CFont m_fontLabel;
	CFont m_fontEnvelope;
	CFont m_fontList;
	int m_firstLabel;
	int m_lastLabel;
	int m_labelLeftMargin;   // distance from the left edge of the page to the left edge of the first label
	int m_labelTopMargin;   // distance from the top of the page to the top edge of the first label
	int m_labelWidth;   // width of the label
	int m_labelHeight;   // height of the label
	int m_labelColumnWidth;  // width of the label plus border
	int m_labelDimension[2];  // number of labels [across page, top to bottom]
	BOOL m_return;
	BOOL m_lastNameFirst;
	int m_columnWidth;
	int m_margin[4];  // left, top, right, bottom
	short m_paperSizeBook;
	short m_paperOrientationBook;
	short m_paperSizeLabel;
	short m_paperOrientationLabel;
	short m_paperSizeEnvelope;
	short m_paperOrientationEnvelope;
	short m_paperSizeList;
	short m_paperOrientationList;

protected:
	CSize m_sizeDoc;

// Generated message map functions
protected:
	//{{AFX_MSG(CAddressDoc)
	afx_msg void OnWriteUnixMailrc();
	afx_msg void OnUpdateFilePrint(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFilePrintPreview(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFilePrintSetup(CCmdUI* pCmdUI);
	afx_msg void OnFileExportCsv();
	afx_msg void OnFileExportPalmCsv();
	afx_msg void OnFileMerge();
	afx_msg void OnFileWriteSelectedToFile();
	afx_msg void OnUpdateFileWriteSelectedToFile(CCmdUI* pCmdUI);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CAddressDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};


class CAddressTable: public CObject
{
public:
	CAddressTable(void)
	{
		int j;
		m_prev = NULL;
		m_next = NULL;
		m_numberOfLinesBook = 0;
		m_numberOfLinesLabel = 0;
		m_selected = FALSE;
		m_drawCoord[0] = 0;
		m_drawCoord[1] = 0;
		m_drawCoord[2] = 0;
		m_drawCoord[3] = 0;
		m_name[0] = '\0';
		m_nameLastFirst[0] = '\0';
		for (j = 0; j < 10; j++)
		{
			m_email[j][0] = '\0';
			m_alias[j][0] = '\0';
			m_address[j][0] = '\0';
			m_phone[j][0] = '\0';
			m_note[j][0] = '\0';
		}
		for (j = 0; j < 100; j++)
		{
			m_listLineBreak[j] = NULL;
		}
	};
	~CAddressTable();

// Attributes
public:
	char m_alias[10][256];
	char m_name[256];
	char m_nameLastFirst[256];
	char m_address[10][256];
	char m_email[10][256];
	char m_phone[10][256];
	char m_note[10][256];
	int m_numberOfLinesBook;
	int m_numberOfLinesLabel;
	BOOL m_selected;
	char m_list[4096];
	char *m_listLineBreak[100];   // pointers to chars where line breaks occur in list view
	int m_drawCoord[4];  // upper left and lower right coords of drawing bounding box

	char m_workcellName[256];

	CAddressTable *m_prev;
	CAddressTable *m_next;

// Operations
public:
	BOOL rpInsertTable(CAddressTable **, BOOL);
	BOOL rpMergeEntry(CAddressTable **);
	void rpNameLastFirst(void);
};


class CAddressAlias: public CObject
{
public:
	CAddressAlias(void)
	{
		m_prev = NULL;
		m_next = NULL;
	};
	~CAddressAlias();

// Attributes
public:
	char m_alias[256];
	char m_email[256];
	CAddressAlias *m_prev;
	CAddressAlias *m_next;

// Operations
public:
	BOOL rpInsertAlias(CAddressAlias **);
};


class CAddressGroup: public CObject
{
public:
	CAddressGroup(void)
	{
		m_prev = NULL;
		m_next = NULL;
	};
	~CAddressGroup();

// Attributes
public:
	char m_groupName[256];
	CAddressTable *m_ptrTable;
	CAddressGroup *m_prev;
	CAddressGroup *m_next;

// Operations
public:
	BOOL rpInsertGroup(CAddressGroup **, BOOL);
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESSDOC_H__41A4E9BA_D2D4_4DA9_B764_21F314A2A46E__INCLUDED_)
