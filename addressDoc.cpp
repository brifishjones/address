// addressDoc.cpp : implementation of the CAddressDoc class
//

#include "stdafx.h"
#include "address.h"

#include "addressDoc.h"
#include "addressView.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressDoc

IMPLEMENT_DYNCREATE(CAddressDoc, CDocument)

BEGIN_MESSAGE_MAP(CAddressDoc, CDocument)
	//{{AFX_MSG_MAP(CAddressDoc)
	ON_COMMAND(ID_FILE_EXPORT_MAILRC, OnWriteUnixMailrc)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT, OnUpdateFilePrint)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_PREVIEW, OnUpdateFilePrintPreview)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_SETUP, OnUpdateFilePrintSetup)
	ON_COMMAND(ID_FILE_EXPORT_CSV, OnFileExportCsv)
	ON_COMMAND(ID_FILE_EXPORT_PALM_CSV, OnFileExportPalmCsv)
	ON_COMMAND(ID_FILE_MERGE, OnFileMerge)
	ON_COMMAND(ID_FILE_WRITE_SELECTED_TO_FILE, OnFileWriteSelectedToFile)
	ON_UPDATE_COMMAND_UI(ID_FILE_WRITE_SELECTED_TO_FILE, OnUpdateFileWriteSelectedToFile)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CAddressDoc, CDocument)
	//{{AFX_DISPATCH_MAP(CAddressDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//      DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IAddress to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {1EAB1A87-CAFD-4A8A-B43B-77BBDBE7EE49}
static const IID IID_IAddress =
{ 0x1eab1a87, 0xcafd, 0x4a8a, { 0xb4, 0x3b, 0x77, 0xbb, 0xdb, 0xe7, 0xee, 0x49 } };

BEGIN_INTERFACE_MAP(CAddressDoc, CDocument)
	INTERFACE_PART(CAddressDoc, IID_IAddress, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressDoc construction/destruction

CAddressDoc::CAddressDoc()
{
	// TODO: add one-time construction code here

	EnableAutomation();

	AfxOleLockApp();

	m_ptrTableList = NULL;
	m_ptrAliasList = NULL;
	m_ptrGroupList = NULL;
	m_groupSelected = NULL;
	m_returnAddress[0] = '\0';
	m_lastNameFirst = FALSE;

	m_offsetBorder = 0;

	m_lineSpacing = 14;
	m_lineSpacingList = 14;
	m_lineSpacingEnvelope = 16;
	m_returnSizeEnvelope = 100;
	m_firstLabel = 0;
	m_lastLabel = 0;

	m_return = FALSE;
	m_columnWidth = 240;
	m_pageWidth = 850;
	m_pageHeight = 1100;
	m_printableWidth = 800;
	m_printableHeight = 1047;
	m_physicalOffsetX = 25;
	m_physicalOffsetY = 7;
	m_fontBook.CreatePointFont(90, "Times New Roman", NULL);
	m_fontList.CreatePointFont(90, "Times New Roman", NULL);
	m_fontLabel.CreatePointFont(120, "Times New Roman", NULL);
	m_fontEnvelope.CreatePointFont(120, "Times New Roman", NULL);

	m_margin[0] = 50;
	m_margin[1] = 50;
	m_margin[2] = 50;
	m_margin[3] = 50;

	m_cellSpacing[0] = 100;
	m_cellSpacing[1] =  50;

	m_tableSize[0] = m_cellSpacing[0] * 10;
	m_tableSize[1] = m_cellSpacing[1];

	m_sizeDoc = 0;

	// initialize paper and envelope sizes based on print setup
	CDC dc;
	CPrintDialog dlg(FALSE);
	CWinApp* app = AfxGetApp();

	app->GetPrinterDeviceDefaults(&dlg.m_pd);

	LPDEVMODE lp = (LPDEVMODE) ::GlobalLock(dlg.m_pd.hDevMode);
	ASSERT(lp);

	if (lp->dmPaperSize == DMPAPER_A4)
	{
		m_paperSizeBook = DMPAPER_A4;
		m_paperOrientationBook = DMORIENT_PORTRAIT;
		m_paperSizeList = DMPAPER_A4;
		m_paperOrientationList = DMORIENT_PORTRAIT;
		m_paperSizeLabel = DMPAPER_A4;
		m_paperOrientationLabel = DMORIENT_PORTRAIT;
		m_paperSizeEnvelope = DMPAPER_ENV_DL;
		m_paperOrientationEnvelope = DMORIENT_LANDSCAPE;

		m_labelLeftMargin = 13;
		m_labelTopMargin = 51;
		m_labelWidth = 252;
		m_labelHeight = 133;
		m_labelColumnWidth = 270;
		m_labelDimension[0] = 3;
		m_labelDimension[1] = 8;
	}
	else
	{
		m_paperSizeBook = DMPAPER_LETTER;
		m_paperOrientationBook = DMORIENT_PORTRAIT;
		m_paperSizeList = DMPAPER_LETTER;
		m_paperOrientationList = DMORIENT_PORTRAIT;
		m_paperSizeLabel = DMPAPER_LETTER;
		m_paperOrientationLabel = DMORIENT_PORTRAIT;
		m_paperSizeEnvelope = DMPAPER_ENV_10;
		m_paperOrientationEnvelope = DMORIENT_LANDSCAPE;

		m_labelLeftMargin = 12;
		m_labelTopMargin = 50;
		m_labelWidth = 262;
		m_labelHeight = 100;
		m_labelColumnWidth = 278;
		m_labelDimension[0] = 3;
		m_labelDimension[1] = 10;

	}

	::GlobalUnlock(dlg.m_pd.hDevMode);

	dlg.CreatePrinterDC();
	dc.Attach(dlg.m_pd.hDC);

}

CAddressDoc::~CAddressDoc()
{
	//CWinApp* app = AfxGetApp();
	//if (((CAddressApp *)app)->m_temporaryKey[0] != '\0')
	//{
	//	((CAddressApp *)app)->rpWriteRegistry();
	//}

	rpFreeDoc();
	TRACE1("~CAddressDoc() IsModified(%d)\n", IsModified());
	AfxOleUnlockApp();
}

BOOL CAddressDoc::OnNewDocument()
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)
	CWnd *pCBox;
	char s[32];

	rpFreeDoc();

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	sprintf(s, "%d", m_lineSpacing);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	sprintf(s, "%d", m_columnWidth);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	m_return = FALSE;
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
	pCBox->ShowWindow(SW_HIDE);

	rpUpdateGroupMenu();
	
	UpdateAllViews(NULL, (LPARAM)0, NULL);
	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);
	pView->RedrawWindow();

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CAddressDoc serialization

void CAddressDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CAddressDoc diagnostics

#ifdef _DEBUG
void CAddressDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CAddressDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CAddressDoc commands

CAddressTable::~CAddressTable()
{

}


CAddressAlias::~CAddressAlias()
{

}

CAddressGroup::~CAddressGroup()
{

}


BOOL CAddressTable::rpMergeEntry(CAddressTable **ptrTable)
// merge this table entry with ptrTable then delete this
// NOT USED as of 23 Oct 2004.  Duplicate entries no longer possible
{
	int i;
	int j;
	int k;
	int n;
	char *t;
	BOOL dup;
	CAddressTable *curTable;
	//CAddressGroup *ptrGroup;
	CAddressGroup *prevGroup = NULL;

	curTable = *ptrTable;
	curTable->m_selected = m_selected;

	// address
	for (i = 0; i < 9 && curTable->m_address[i][0] != '\0'; i++);
	k = 0;
	for (j = i; j < 9 && m_address[k][0] != '\0'; j++)
	{
		dup = FALSE;
		for (n = 0; n < i; n++)
		{
			if (strcmp(curTable->m_address[n], m_address[k]) == 0)
			{
				dup = TRUE;
			}
		}
		if (dup == FALSE)
		{
			sprintf(curTable->m_address[j], "%s", m_address[k]);

			for (t = m_address[k]; *t != '\0'; t++)
			{
				if (*t == ';')
				{
					curTable->m_numberOfLinesBook++;
					curTable->m_numberOfLinesLabel++;
				}
			}
			curTable->m_numberOfLinesBook++;
			curTable->m_numberOfLinesLabel++;
		}
		else
		{
			j--;
		}
		k++;
	}

	// email
	for (i = 0; i < 9 && curTable->m_email[i][0] != '\0'; i++);
	k = 0;
	for (j = i; j < 9 && m_email[k][0] != '\0'; j++)
	{
		dup = FALSE;
		for (n = 0; n < i; n++)
		{
			if (strcmp(curTable->m_email[n], m_email[k]) == 0)
			{
				dup = TRUE;
			}
		}
		if (dup == FALSE)
		{
			sprintf(curTable->m_email[j], "%s", m_email[k]);
			sprintf(curTable->m_alias[j], "%s", m_alias[k]);
			curTable->m_numberOfLinesBook++;
		}
		else
		{
			j--;
		}
		k++;
	}

	// phone
	for (i = 0; i < 9 && curTable->m_phone[i][0] != '\0'; i++);
	k = 0;
	for (j = i; j < 9 && m_phone[k][0] != '\0'; j++)
	{
		dup = FALSE;
		for (n = 0; n < i; n++)
		{
			if (strcmp(curTable->m_phone[n], m_phone[k]) == 0)
			{
				dup = TRUE;
			}
		}
		if (dup == FALSE)
		{
			sprintf(curTable->m_phone[j], "%s", m_phone[k]);
			curTable->m_numberOfLinesBook++;
		}
		else
		{
			j--;
		}
		k++;
	}

	// note
	for (i = 0; i < 9 && curTable->m_note[i][0] != '\0'; i++);
	k = 0;
	for (j = i; j < 9 && m_note[k][0] != '\0'; j++)
	{
		dup = FALSE;
		for (n = 0; n < i; n++)
		{
			if (strcmp(curTable->m_note[n], m_note[k]) == 0)
			{
				dup = TRUE;
			}
		}
		if (dup == FALSE)
		{
			sprintf(curTable->m_note[j], "%s", m_note[k]);
			curTable->m_numberOfLinesBook++;
		}
		else
		{
			j--;
		}
		k++;
	}

	/**
	// remove entries from group list for table entry (this) before it is deleted
	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL;)
	{
		if (ptrGroup->m_ptrTable == this)
		{
			if (ptrGroup == pDoc->m_groupSelected)
			{
				pDoc->m_groupSelected = ptrGroup->m_next;
			}
			if (ptrGroup == pDoc->m_ptrGroupList)
			// first in list
			{
				pDoc->m_ptrGroupList = ptrGroup->m_next;
			}
			if (ptrGroup->m_next != NULL)
			{
				ptrGroup->m_next->m_prev = ptrGroup->m_prev;
			}
			if (ptrGroup->m_prev != NULL)
			{
				ptrGroup->m_prev->m_next = ptrGroup->m_next;
			}
			prevGroup = ptrGroup;
			ptrGroup = ptrGroup->m_next;
			delete prevGroup;
		}
		else
		{
			ptrGroup = ptrGroup->m_next;
		}
	}
**/


	delete this;
	return TRUE;

}


void CAddressTable::rpNameLastFirst(void)
// sets m_nameLastFirst based on m_name
{
	char s[256];
	char *t;
	char *u;
	char firstName[256];
	char lastName[256];

	if (m_name[0] == '\0')
	{
		m_nameLastFirst[0] = '\0';
		return;
	}

	sprintf(s, "%s", m_name);
	u = s;
	for (t = s; *t != '\0'; t++); // go to the end of name

	for (; *t != ' ' && t != s; t--);
	if (*t == ' ')
	{
		*t = '\0';
		t++;
		sprintf(firstName, "%s", u);
		sprintf(lastName, "%s", t);
	}
	else  // one name entry
	{
		sprintf(firstName, "%s", u);
		lastName[0] = '\0';
	}

	if (lastName[0] != '\0')
	{
		sprintf(m_nameLastFirst, "%s, %s", lastName, firstName);
	}
	else
	{
		sprintf(m_nameLastFirst, "%s", firstName);
	}

	return;
}

BOOL CAddressTable::rpInsertTable(CAddressTable **ptrTableList, BOOL lastNameFirst)
// Put this table at the head, tail, or in the middle of ptrTableList, which
// is m_ptrTableList in CAddressDoc class.
{
	CAddressTable *prevTable = NULL;
	CAddressTable *curTable;

	rpNameLastFirst();

	if (*ptrTableList == NULL)
	{
		*ptrTableList = this;
		return TRUE;
	}

	if (lastNameFirst == FALSE)
	{
		for (curTable = *ptrTableList;
			curTable != NULL && strcmp(m_name, curTable->m_name) > 0;
			curTable = curTable->m_next)
		{
			prevTable = curTable;
		}

		// if duplicate merge with curTable
		//if (curTable != NULL && strcmp(m_name, curTable->m_name) == 0)
		//{
		//	rpMergeEntry(&curTable);
		//	return TRUE;
		//}
	}
	else
	{
		for (curTable = *ptrTableList;
			curTable != NULL && strcmp(m_nameLastFirst, curTable->m_nameLastFirst) > 0;
			curTable = curTable->m_next)
		{
			prevTable = curTable;
		}

		// if duplicate merge with curTable
		//if (curTable != NULL && strcmp(m_nameLastFirst, curTable->m_nameLastFirst) == 0)
		//{
		//	rpMergeEntry(&curTable);
		//	return TRUE;
		//}
	}

	if (prevTable == NULL)
	{
		// insert at beginning of list
		*ptrTableList = this;
		this->m_next = curTable;
		curTable->m_prev = this;
	}
	else if (curTable == NULL)
	{
		// insert at end of list
		prevTable->m_next = this;
		this->m_prev = prevTable;
	}
	else
	{
		// insert in middle of list
		prevTable->m_next = this;
		this->m_prev = prevTable;
		curTable->m_prev = this;
		this->m_next = curTable;
	}

	return TRUE;
}

BOOL CAddressAlias::rpInsertAlias(CAddressAlias **ptrAliasList)
// Put this alias at the head, tail, or in the middle of ptrAliasList, which
// is m_ptrAliasList in CMainFrame class.
{

	CAddressAlias *prevAlias = NULL;
	CAddressAlias *curAlias;
	char s[512];
	char t[512];

	if (*ptrAliasList == NULL)
	{
		*ptrAliasList = this;
		return TRUE;
	}

	sprintf(s, "%s %s", m_alias, m_email);
	sprintf(t, "%s %s", (*ptrAliasList)->m_alias, (*ptrAliasList)->m_email);
	for (curAlias = *ptrAliasList;
		curAlias != NULL && strcmp(s, t) > 0;
		curAlias = curAlias->m_next)
	{
		prevAlias = curAlias;
		if (curAlias->m_next != NULL)
		{
			sprintf(t, "%s %s", curAlias->m_next->m_alias, curAlias->m_next->m_email);
		}
	}

	if (prevAlias == NULL)
	{
		// insert at beginning of list
		*ptrAliasList = this;
		this->m_next = curAlias;
		curAlias->m_prev = this;
	}
	else if (curAlias == NULL)
	{
		// insert at end of list
		prevAlias->m_next = this;
		this->m_prev = prevAlias;
	}
	else
	{
		// insert in middle of list
		prevAlias->m_next = this;
		this->m_prev = prevAlias;
		curAlias->m_prev = this;
		this->m_next = curAlias;
	}

	return TRUE;
}


BOOL CAddressGroup::rpInsertGroup(CAddressGroup **ptrGroupList, BOOL lastNameFirst)
// Put this group at the head, tail, or in the middle of ptrGroupList, which
// is m_ptrGroupList in CMainFrame class.
{
	CAddressGroup *prevGroup = NULL;
	CAddressGroup *curGroup;
	char s[512];
	char t[512];

	if (*ptrGroupList == NULL)
	{
		*ptrGroupList = this;
		return TRUE;
	}

	if (lastNameFirst == FALSE)
	{
		sprintf(s, "%s %s", m_groupName, m_ptrTable->m_name);
		sprintf(t, "%s %s", (*ptrGroupList)->m_groupName, (*ptrGroupList)->m_ptrTable->m_name);
		for (curGroup = *ptrGroupList;
			curGroup != NULL && strcmp(s, t) > 0;
			curGroup = curGroup->m_next)
		{
			prevGroup = curGroup;
			if (curGroup->m_next != NULL)
			{
				sprintf(t, "%s %s", curGroup->m_next->m_groupName, curGroup->m_next->m_ptrTable->m_name);
			}
		}
	}
	else
	{
		sprintf(s, "%s %s", m_groupName, m_ptrTable->m_nameLastFirst);
		sprintf(t, "%s %s", (*ptrGroupList)->m_groupName, (*ptrGroupList)->m_ptrTable->m_nameLastFirst);
		for (curGroup = *ptrGroupList;
			curGroup != NULL && strcmp(s, t) > 0;
			curGroup = curGroup->m_next)
		{
			prevGroup = curGroup;
			if (curGroup->m_next != NULL)
			{
				sprintf(t, "%s %s", curGroup->m_next->m_groupName, curGroup->m_next->m_ptrTable->m_nameLastFirst);
			}
		}
	}

	if (prevGroup == NULL)
	{
		// insert at beginning of list
		*ptrGroupList = this;
		this->m_next = curGroup;
		curGroup->m_prev = this;
	}
	else if (curGroup == NULL)
	{
		// insert at end of list
		prevGroup->m_next = this;
		this->m_prev = prevGroup;
	}
	else
	{
		// insert in middle of list
		prevGroup->m_next = this;
		this->m_prev = prevGroup;
		curGroup->m_prev = this;
		this->m_next = curGroup;
	}

	return TRUE;
}



BOOL CAddressDoc::rpWriteDoc(LPCTSTR lpszPathName)
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}

	FILE *fp;
	CAddressTable *ptrTable = NULL;
	int j;
	LOGFONT lf;
	CDC *pDC;
	pDC = pView->GetDC( );
	CAddressGroup *ptrGroup;

	fp = fopen((char *)lpszPathName, "w");

	if (fp == NULL)
	{
		return FALSE;
	}

	fprintf(fp, "~¡¢£¤¥¦§¨©ª«­®¯°±²³´µ·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ\n\n");

	for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (ptrTable != m_ptrTableList)
		{
			fprintf(fp, "\n");
		}	
		fprintf(fp, "name %s\n", ptrTable->m_name);

		fprintf(fp, "selected %d\n", ptrTable->m_selected);

		for (j = 0; j < 9 && ptrTable->m_address[j][0] != '\0'; j++)
		{
			fprintf(fp, "address %s\n", ptrTable->m_address[j]);
		}

		for (j = 0; j < 9 && ptrTable->m_email[j][0] != '\0'; j++)
		{
			fprintf(fp, "email %s", ptrTable->m_email[j]);
			//if (ptrTable->m_alias[j][0] != '\0')
			//{
			//	fprintf(fp, " %s\n", ptrTable->m_alias[j]);
			//}
			if (ptrTable->m_alias[j][0] != '\0')
			{
				fprintf(fp, " %s\n", ptrTable->m_alias[j]);
			}
			else
			{
				fprintf(fp, "\n");
			}
		}

		for (j = 0; j < 9 && ptrTable->m_phone[j][0] != '\0'; j++)
		{
			fprintf(fp, "phone %s\n", ptrTable->m_phone[j]);
		}

		for (j = 0; j < 9 && ptrTable->m_note[j][0] != '\0'; j++)
		{
			fprintf(fp, "note %s\n", ptrTable->m_note[j]);
		}

	}

	if (m_returnAddress[0] != '\0')
	{
		fprintf(fp, "\nreturn %s\n", m_returnAddress);
	}

	fprintf(fp, "\nlastNameFirst %d\n", m_lastNameFirst);

	m_fontBook.GetLogFont(&lf);
	j = -MulDiv(lf.lfHeight, 72, pDC->GetDeviceCaps(LOGPIXELSY));
	fprintf(fp, "\nbook \"%s\" %d %d %d\n", lf.lfFaceName, j, m_lineSpacing, m_columnWidth);

	m_fontList.GetLogFont(&lf);
	j = -MulDiv(lf.lfHeight, 72, pDC->GetDeviceCaps(LOGPIXELSY));
	fprintf(fp, "list \"%s\" %d %d\n", lf.lfFaceName, j, m_lineSpacingList);

	m_fontEnvelope.GetLogFont(&lf);
	j = -MulDiv(lf.lfHeight, 72, pDC->GetDeviceCaps(LOGPIXELSY));
	fprintf(fp, "envelope \"%s\" %d %d %d\n", lf.lfFaceName, j, m_lineSpacingEnvelope, m_returnSizeEnvelope);

	m_fontLabel.GetLogFont(&lf);
	j = -MulDiv(lf.lfHeight, 72, pDC->GetDeviceCaps(LOGPIXELSY));
	fprintf(fp, "label \"%s\" %d", lf.lfFaceName, j);

	for (ptrGroup = m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == m_ptrGroupList)
		{
			fprintf(fp, "\n\n");
		}
		fprintf(fp, "group %s %s\n", ptrGroup->m_groupName, ptrGroup->m_ptrTable->m_name);
		
		if (ptrGroup->m_next != NULL
			&& strcmp(ptrGroup->m_groupName, ptrGroup->m_next->m_groupName) != 0)
		{
			fprintf(fp, "\n");
		}
	}

	fclose(fp);
	return TRUE;

}


BOOL CAddressDoc::rpUpdateAddressList()
// fill name list in left dialog listbox with names from m_ptrTableList
// selection of names is done in rpComputeDrawPositions()
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CAddressTable *ptrTable = NULL;
	CWnd *pCBox;

	pView->rpLeftDlgListbox();

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
	((CListBox *)pCBox)->ResetContent();

	for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (m_lastNameFirst == FALSE)
		{
			((CListBox *)pCBox)->AddString((ptrTable->m_name));
		}
		else
		{
			((CListBox *)pCBox)->AddString((ptrTable->m_nameLastFirst));
		}
	}

	return TRUE;
}


BOOL CAddressDoc::rpReadDoc(FILE *fp)
// Read an address document file.
{
	CAddressTable *ptrTable = NULL;
	CAddressTable *curTable = NULL;
	CAddressTable *prevTable = NULL;
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	int i;
	int j;
	char s[256];
	char *t;
	int ch = 0;
	char command[32];
	char fontName[64];
	CFont fontBook;
	CFont fontList;
	CFont fontLabel;
	CFont fontEnvelope;
	LOGFONT lf;
	HFONT hfont;
	char groupName[256];
	char memberName[256];
	BOOL done = FALSE;
	BOOL validCommandEncountered = FALSE;
	int numberOfGroups;

	//while(!done)
	while(!done)
	{
		// Read in single line from file
		memset(command, 0, 32); 
		memset(s, 0, 256);
		for (i = 0; (i < 255) && ((ch = fgetc(fp)) != EOF)
			&& (ch != '\n'); i++)
		{
			s[i] = (char)ch;
			if (s[i] == '\t')
			{
				s[i] = ' ';
			}
			//TRACE2("%d: %c\n", ch, s[i]);
		}
		s[i] = (char)ch;
		//TRACE2("%d: %c\n", ch, s[i]);
		if (s[i] == '\n')
		{
			s[i] = '\0';
		}

		sscanf(s, "%s", command);

		if (ch == EOF)
		{
			fclose(fp);
			done = TRUE;
		}

		//else if (strcmp(command, "alias") == 0)
		//{
		//	i = sscanf(s, "%*s %s", 
		//			ptrTable->m_alias[0]);
		//	ptrTable->m_numberOfLinesBook++;
		//	ptrTable->m_numberOfLinesLabel++;
		//
		//}

		//else if (strcmp(command, "First") == 0
		//	|| strcmp(command, "Last") == 0
		//	|| strcmp(command, "Name") == 0)
		else if ((s[0] == 'F' && s[1] == 'i' && s[2] == 'r' && s[3] == 's' && s[4] == 't')
			|| (s[0] == 'L' && s[1] == 'a' && s[2] == 's' && s[3] == 't')
			|| (s[0] == 'N' && s[1] == 'a' && s[2] == 'm' && s[3] == 'e'))
		// test for Outlook Express comma separated file format
		{
			return (rpReadCsvFile(fp));
		}

		else if (s[0] == '"')
		// test for Palm comma separated file format
		{
			return (rpReadPalmCsvFile(fp));
		}

		else if (strcmp(command, "name") == 0)
		{
			validCommandEncountered = TRUE;
			// insert previous entry in list
			if (ptrTable != NULL)
			{
				ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
			}
			//if (ptrTable->m_name[0] == '\0')
			//{
			//	sprintf(ptrTable->m_name, "%s", ptrTable->m_alias[0]);
			//	ptrTable->m_name[0] = toupper(ptrTable->m_name[0]);
			//}
			/**
			if (ptrTable != NULL)
			{
				if (m_ptrTableList == NULL)
				{
					m_ptrTableList = ptrTable;
				}
				else
				{
					for (curTable = m_ptrTableList;
						curTable != NULL && strcmp(ptrTable->m_name, curTable->m_name) > 0;
						curTable = curTable->m_next)
					{
						prevTable = curTable;
					}

					if (curTable == m_ptrTableList)
					{
						// insert at head of list
						curTable->m_prev = ptrTable;
						ptrTable->m_next = curTable;
						m_ptrTableList = ptrTable;
					}
					else if (curTable != NULL)
					{
						// insert in middle of list
						curTable->m_prev->m_next = ptrTable;
						ptrTable->m_prev = curTable->m_prev;
						curTable->m_prev = ptrTable;
						ptrTable->m_next = curTable;
					}
					else // curTable == NULL
					{
						// last item in list
						prevTable->m_next = ptrTable;
						ptrTable->m_prev = prevTable;
					}
				}
			}
			**/

			ptrTable = new CAddressTable;

			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' '; t++);  // trailing white space after command
			i = 0;
			for (; *t != '\0'; t++)
			{
				ptrTable->m_name[i] = *t;
				i++;
			}
			ptrTable->m_name[i] = '\0';
			ptrTable->m_numberOfLinesBook++;
			ptrTable->m_numberOfLinesLabel++;
		}

		else if (strcmp(command, "selected") == 0 && ptrTable != NULL)
		{
			validCommandEncountered = TRUE;
			sscanf(s, "%*s %d", &ptrTable->m_selected);
		}

		else if (strcmp(command, "email") == 0  && ptrTable != NULL)
		{
			validCommandEncountered = TRUE;
			//for (t = s; *t == ' '; t++);  // leading white space before command
			//for (; *t != ' '; t++);  // command
			//for (; *t == ' '; t++);  // trailing white space after command
			for (j = 0; j < 9 && ptrTable->m_email[j][0] != '\0'; j++);
			//i = 0;
			//for (; *t != '\0' && *t != ' '; t++)
			//{
			//	ptrTable->m_email[j][i] = *t;
			//	i++;
			//}
			//ptrTable->m_email[j][i] = '\0';

			i = sscanf(s, "%*s %s %s", ptrTable->m_email[j], ptrTable->m_alias[j]);
			if (i != 2)
			{
				ptrTable->m_alias[j][0] = '\0';
			}

			ptrTable->m_numberOfLinesBook++;
		}

		else if (strcmp(command, "address") == 0  && ptrTable != NULL)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' '; t++);  // trailing white space after command
			i = 0;
			for (j = 0; j < 9 && ptrTable->m_address[j][0] != '\0'; j++);
			for (; *t != '\0'; t++)
			{
				ptrTable->m_address[j][i] = *t;
				if (ptrTable->m_address[j][i] == ';')
				{
					ptrTable->m_numberOfLinesBook++;
					if (j == 0)
					{
						ptrTable->m_numberOfLinesLabel++;
					}
					if (*(t + 1) != ' ')
					{
						i++;
						ptrTable->m_address[j][i] = ' ';
					}
				}
				i++;
			}
			ptrTable->m_address[j][i] = '\0';
			ptrTable->m_numberOfLinesBook++;
			if (j == 0)
			{
				ptrTable->m_numberOfLinesLabel++;
			}
		}

		else if (strcmp(command, "phone") == 0  && ptrTable != NULL)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' '; t++);  // trailing white space after command
			for (j = 0; j < 9 && ptrTable->m_phone[j][0] != '\0'; j++);
			i = 0;
			for (; *t != '\0'; t++)
			{
				ptrTable->m_phone[j][i] = *t;
				i++;
			}
			ptrTable->m_phone[j][i] = '\0';
			ptrTable->m_numberOfLinesBook++;
		}

		else if (strcmp(command, "note") == 0  && ptrTable != NULL)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' '; t++);  // trailing white space after command
			for (j = 0; j < 9 && ptrTable->m_note[j][0] != '\0'; j++);
			i = 0;
			for (; *t != '\0'; t++)
			{
				ptrTable->m_note[j][i] = *t;
				i++;
			}
			ptrTable->m_note[j][i] = '\0';
			ptrTable->m_numberOfLinesBook++;
		}

		else if (strcmp(command, "lastNameFirst") == 0)
		{
			validCommandEncountered = TRUE;
			i = sscanf(s, "%*s %d", &m_lastNameFirst);
			//if (i == 1)
			//{
			//	rpSortName();
			//}
		}

		else if (strcmp(command, "return") == 0)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' '; t++);  // trailing white space after command
			i = 0;
			for (; *t != '\0'; t++)
			{
				m_returnAddress[i] = *t;
				i++;
			}
			m_returnAddress[i] = '\0';
		}

		else if (strcmp(command, "group") == 0)
		{
			validCommandEncountered = TRUE;
			i = sscanf(s, "%*s %s", groupName);

			// if new groupName make certain maximum number of groups is not exceeded
			numberOfGroups = 0;
			for (ptrGroup = m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
			{
				if (ptrGroup == m_ptrGroupList
					|| (ptrGroup->m_prev != NULL
					&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
				{
					if (strcmp(ptrGroup->m_groupName, groupName) != 0)
					{
						numberOfGroups++;
					}
				}
			}

			if (i == 1 && numberOfGroups <= MAXIMUM_NUMBER_GROUPS)
			{
				ptrGroup = new CAddressGroup;

				for (t = s; *t == ' '; t++);  // leading white space before command
				for (; *t != ' '; t++);  // command
				for (; *t == ' '; t++);  // trailing white space after command
				for (; *t != ' '; t++);  // groupName
				for (; *t == ' '; t++);  // trailing white space after groupName
				i = 0;
				for (; *t != '\0'; t++)
				{
					memberName[i] = *t;
					i++;
				}
				memberName[i] = '\0';

				sprintf(ptrGroup->m_groupName, "%s", groupName);

				// match group member name to name in ptrTableList
				for (curTable = m_ptrTableList;
					curTable != NULL && (strcmp(curTable->m_name, memberName) != 0);
					curTable = curTable->m_next);

				if (curTable != NULL)
				{
					ptrGroup->m_ptrTable = curTable;
					ptrGroup->rpInsertGroup(&m_ptrGroupList, m_lastNameFirst);
				}
				else
				{
					delete ptrGroup;
				}
			}
		}

		else if (strcmp(command, "book") == 0)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' ' || *t == '"'; t++);  // trailing white space and double quote after command
			i = 0;
			for (; *t != '\0' && *t != '"'; t++)
			{
				fontName[i] = *t;
				i++;
			}
			fontName[i] = '\0';
			t++;
			i = sscanf(t, "%d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				fontBook.CreatePointFont(j * 10, fontName, NULL);
				fontBook.GetLogFont(&lf);
				hfont = CreateFontIndirect(&lf); 
				m_fontBook.m_hObject = hfont;
			}
			i = sscanf(t, "%*d %d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				m_lineSpacing = j;
			}
			i = sscanf(t, "%*d %*d %d", &j);
			if (i == 1)
			{
				j = MAX(j, 100);
				j = MIN(j, m_printableWidth);
				m_columnWidth = j;
			}
		}

		else if (strcmp(command, "list") == 0)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' ' || *t == '"'; t++);  // trailing white space and double quote after command
			i = 0;
			for (; *t != '\0' && *t != '"'; t++)
			{
				fontName[i] = *t;
				i++;
			}
			fontName[i] = '\0';
			t++;
			i = sscanf(t, "%d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				fontList.CreatePointFont(j * 10, fontName, NULL);
				fontList.GetLogFont(&lf);
				hfont = CreateFontIndirect(&lf); 
				m_fontList.m_hObject = hfont;
			}
			i = sscanf(t, "%*d %d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				m_lineSpacingList = j;
			}
		}
		else if (strcmp(command, "envelope") == 0)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' ' || *t == '"'; t++);  // trailing white space and double quote after command
			i = 0;
			for (; *t != '\0' && *t != '"'; t++)
			{
				fontName[i] = *t;
				i++;
			}
			fontName[i] = '\0';
			t++;
			i = sscanf(t, "%d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				fontEnvelope.CreatePointFont(j * 10, fontName, NULL);
				fontEnvelope.GetLogFont(&lf);
				hfont = CreateFontIndirect(&lf); 
				m_fontEnvelope.m_hObject = hfont;
			}
			i = sscanf(t, "%*d %d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				m_lineSpacingEnvelope = j;
			}
			i = sscanf(t, "%*d %*d %d", &j);
			if (i == 1)
			{
				j = MAX(j, 20);
				j = MIN(j, 100);
				m_returnSizeEnvelope = j;
			}
		}

		else if (strcmp(command, "label") == 0)
		{
			validCommandEncountered = TRUE;
			for (t = s; *t == ' '; t++);  // leading white space before command
			for (; *t != ' '; t++);  // command
			for (; *t == ' ' || *t == '"'; t++);  // trailing white space and double quote after command
			i = 0;
			for (; *t != '\0' && *t != '"'; t++)
			{
				fontName[i] = *t;
				i++;
			}
			fontName[i] = '\0';
			t++;
			i = sscanf(t, "%d", &j);
			if (i == 1)
			{
				j = MAX(j, 8);
				j = MIN(j, 72);
				fontLabel.CreatePointFont(j * 10, fontName, NULL);
				fontLabel.GetLogFont(&lf);
				hfont = CreateFontIndirect(&lf); 
				m_fontLabel.m_hObject = hfont;
			}
		}


	} // end while (!done)

	// insert final entry in list
	if (ptrTable != NULL)
	{
		ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
		//if (ptrTable->m_name[0] == '\0')
		//{
		//	sprintf(ptrTable->m_name, "%s", ptrTable->m_alias[0]);
		//	ptrTable->m_name[0] = toupper(ptrTable->m_name[0]);
		//}

		/**
		if (m_ptrTableList == NULL)
		{
			m_ptrTableList = ptrTable;
		}
		else
		{
			for (curTable = m_ptrTableList;
				curTable != NULL && strcmp(ptrTable->m_name, curTable->m_name) > 0;
				curTable = curTable->m_next)
			{
				prevTable = curTable;
			}

			if (curTable == m_ptrTableList)
			{
				// insert at head of list
				curTable->m_prev = ptrTable;
				ptrTable->m_next = curTable;
				m_ptrTableList = ptrTable;
			}
			else if (curTable != NULL)
			{
				// insert in middle of list
				curTable->m_prev->m_next = ptrTable;
				ptrTable->m_prev = curTable->m_prev;
				curTable->m_prev = ptrTable;
				ptrTable->m_next = curTable;
			}
			else // curTable == NULL
			{
				// last item in list
				prevTable->m_next = ptrTable;
				ptrTable->m_prev = prevTable;
			}
		}
		***/
	}

	/**
	// create alias list
	for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		for (j = 0; j < 9; j++)
		{
			if (ptrTable->m_alias[j][0] != '\0')
			{
				if (ptrTable->m_email[j][0] != '\0')
				{
					ptrAlias = new CAddressAlias;
					sprintf(ptrAlias->m_alias, "%s", ptrTable->m_alias[j]);
					sprintf(ptrAlias->m_email, "%s", ptrTable->m_email[j]);
					ptrAlias->rpInsertAlias(&m_ptrAliasList);

				}
			}
		}

	}
	**/

	rpUpdateAddressList();


	return validCommandEncountered;
}

BOOL CAddressDoc::rpReadPalmCsvFile(FILE *fp)
// to do:
// test with double quotes in a field
// tabs?
{
	CAddressTable *ptrTable = NULL;
	CAddressTable *curTable = NULL;
	CAddressTable *prevTable = NULL;
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	int i;
	int j;
	int k;
	char s[1024];
	char *t;
	char *u;
	int ch = 0;
	char heading[40][32];
	char entry[40][256];
	CFont fontBook;
	CFont fontList;
	CFont fontLabel;
	CFont fontEnvelope;
	BOOL done = FALSE;
	BOOL openQuote = FALSE;

	rewind(fp);

	sprintf(heading[0], "Last Name");
	sprintf(heading[1], "First Name");
	sprintf(heading[2], "Title");
	sprintf(heading[3], "Company");
	sprintf(heading[4], "Work Number");
	sprintf(heading[5], "Home Number");
	sprintf(heading[6], "Fax");
	sprintf(heading[7], "Other");
	sprintf(heading[8], "Email");
	sprintf(heading[9], "Street");
	sprintf(heading[10], "City");
	sprintf(heading[11], "State");
	sprintf(heading[12], "Zip");
	sprintf(heading[13], "Country");
	//sprintf(heading[14], "Custom 1");
	//sprintf(heading[15], "Custom 2");
	//sprintf(heading[16], "Custom 3");
	//sprintf(heading[17], "Custom 4");
	//sprintf(heading[18], "Note Card");
	//sprintf(heading[19], "Private");
	k = 14;

	// read entries into table
	i = 0;
	j = 0;
	while (TRUE)
	{
		ch = fgetc(fp);
		if (ch == 0x0D)
		// change carriage return to space
		{
			ch = 0x20;
		}
		//TRACE3("%03d: %d %c\n", j, ch, (char)ch);
		entry[i][MIN(j, 255)] = (char)ch;

		if (ch == '"' && openQuote)
		{
			openQuote = FALSE;	
		}
		else if (ch == '"' && !openQuote)
		{
			openQuote = TRUE;
		}

		if ((ch == ',' && !openQuote) || (ch == '\n' && !openQuote) || ch == EOF)
		{
			entry[i][j] = '\0';

			// remove rear quote before comma
			if (entry[i][j - 1] == '"')
			{
				entry[i][j - 1] = '\0';
			}
			// remove quote at beginning of entry string
			if (entry[i][0] == '"')
			{
				sprintf(entry[i], "%s", &entry[i][1]);
			}

			// set double quotes in entry string to single quote
			j = 0;
			for (t = entry[i]; *t != '\0'; t++)
			{
				if (*t != '"' || *(t + 1) != '"')
				{
					entry[i][j] = *t;
					j++;
				}
			}
			entry[i][j] = '\0';

			//rpRemoveQuotes(entry[i]);
			if (i != 9) // not a street address
			{
				for (t = entry[i]; *t == ' ' || *t == 0x0A; t++); // remove leading chars
				j = 0;
				for ( ; *t != '\0'; t++)
				{
					if (*t == '\n')   // bad entry in palm--strip remaining info for field
					{
						*t = '\0';
						t--;
					}
					/**
					if (*t == '\n')
					{
						for (u = t - 1; *u == ' '; u--)  // move back
						{
							j--;
						}
						entry[i][j] = ';';
						entry[i][j + 1] = ' ';
						j += 2;
						for ( ; *(t + 1) == ' ' || *(t + 1) == '\n'; t++);  // move forward
					}
					**/
					else
					{
						entry[i][j] = *t;
						j++;
					}
				}
				entry[i][j] = '\0';
				for (t = &entry[i][j - 1]; *t == ' '; t--)
				{
					*t = '\0';
				}
			}

			//TRACE2("(%d) %s\n", i, entry[i]);
			j = 0;
			i++;
		}
		else
		{
			j++;
		}

		if ((m_ptrTableList == NULL && ((ch == '\n' && !openQuote) || ch == EOF) && i < k) || i >= 40 || j >= 256)
		// illegal file format
		{
			fclose(fp);
			return FALSE;
		}

		if ((ch == '\n' && !openQuote || ch == EOF) && i >= k)
		{

			// create a new table entry
			ptrTable = new CAddressTable;

			for (i = 0; i < k && strcmp(heading[i], "Last Name") != 0; i++);
			for (j = 0; j < k && strcmp(heading[j], "First Name") != 0; j++);
			if (i < k && entry[i][0] != '\0' && j < k && entry[j][0] != '\0')
			// Last Name, First Name
			{
				//rpRemoveQuotes(entry[j]);
				//rpRemoveQuotes(entry[i]);
				sprintf(ptrTable->m_name, "%s %s", entry[j], entry[i]);
			}
			else if (i < k && entry[i][0] != '\0')
			// Last Name
			{
				//rpRemoveQuotes(entry[i]);
				sprintf(ptrTable->m_name, "%s", entry[i]);
			}
			else if (j < k && entry[j][0] != '\0')
			// First Name
			{
				//rpRemoveQuotes(entry[j]);
				sprintf(ptrTable->m_name, "%s", entry[j]);
			}

			ptrTable->m_numberOfLinesBook++;
			ptrTable->m_numberOfLinesLabel++;

			for (i = 0; i < k && strcmp(heading[i], "Email") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// E-mail Address
			{
				//rpRemoveQuotes(entry[i]);
				sprintf(ptrTable->m_email[0], "%s", entry[i]);
				ptrTable->m_numberOfLinesBook++;
			}

			for (i = 0; i < k && strcmp(heading[i], "Home Number") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home Phone
			{
				//rpRemoveQuotes(entry[i]);
				sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				ptrTable->m_numberOfLinesBook++;
			}

			for (i = 0; i < k && strcmp(heading[i], "Work Number") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Work Phone
			{
				//rpRemoveQuotes(entry[i]);
				if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;

			}

			for (i = 0; i < k && strcmp(heading[i], "Fax") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Fax
			{
				//rpRemoveQuotes(entry[i]);
				if (ptrTable->m_phone[1][0] != '\0')
				{
					sprintf(ptrTable->m_phone[2], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
			}

			for (i = 0; i < k && strcmp(heading[i], "Other") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Other phone
			{
				//rpRemoveQuotes(entry[i]);
				if (ptrTable->m_phone[2][0] != '\0')
				{
					sprintf(ptrTable->m_phone[3], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[1][0] != '\0')
				{
					sprintf(ptrTable->m_phone[2], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
			}

			s[0] = '\0';
			for (i = 0; i < k && strcmp(heading[i], "Street") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Street
			{
				//if (entry[i][0] == '"')
				//{
				//	t = entry[i] + 1;
				//}
				//else
				//{
				//	t = entry[i];
				//}
				//s[0] = *t;
				//j = 1;
				//for (t = t + 1; *t != '\0'; t++)
				//rpRemoveQuotes(entry[i]);
				for (t = entry[i]; *t == ' ' || *t == 0x0A; t++); // remove leading chars
				j = 0;
				for ( ; *t != '\0'; t++)
				{
					if (*t == '\n')
					{
						for (u = t - 1; *u == ' '; u--)  // move back
						{
							j--;
						}
						s[j] = ';';
						s[j + 1] = ' ';
						j += 2;
						for ( ; *(t + 1) == ' ' || *(t + 1) == '\n'; t++);  // move forward

					//	s[j] = ';';
					//	j++;
					//	s[j] = ' ';
					//	j++;
					//	while (*(t + 1) == ' ')
					//	{
					//		t++;
					//	}

						ptrTable->m_numberOfLinesBook++;
						ptrTable->m_numberOfLinesLabel++;
					}
					//else if (*t == '"')
					//{
					//	s[j] = '\0';
					//	*t = '\0';
					//}
					else
					{
						s[j] = *t;
						j++;
					}
				}
				s[j] = '\0';
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
				sprintf(ptrTable->m_address[0], "%s", s);	
			}

			for (i = 0; i < k && strcmp(heading[i], "City") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// City
			{
				//rpRemoveQuotes(entry[i]);
				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
				sprintf(ptrTable->m_address[0], "%s", s);
			}
			for (i = 0; i < k && strcmp(heading[i], "State") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// State
			{
				//rpRemoveQuotes(entry[i]);
				if (s[0] != '\0' && entry[i - 1] != '\0')
				{
					sprintf(s, "%s, %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city field is empty
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				sprintf(ptrTable->m_address[0], "%s", s);	
			}
			for (i = 0; i < k && strcmp(heading[i], "Zip") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Postal Code
			{
				//rpRemoveQuotes(entry[i]);
				if (s[0] != '\0' && (entry[i - 2][0] != '\0' || entry[i - 1][0] != '\0'))
				{
					sprintf(s, "%s  %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city and state fields are empty but address field is not
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				sprintf(ptrTable->m_address[0], "%s", s);	
			}

			for (i = 0; i < k && strcmp(heading[i], "Country") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Country
			{
				//rpRemoveQuotes(entry[i]);
				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				sprintf(ptrTable->m_address[0], "%s", s);
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
			}
			
			ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
			i = 0;
			j = 0;
			openQuote = FALSE;
		}

		if (ch == EOF)
		{
			break;
		}		
	}

	rpUpdateAddressList();
	return TRUE;
}



BOOL CAddressDoc::rpReadCsvFile(FILE *fp)
{
	CAddressTable *ptrTable = NULL;
	CAddressTable *curTable = NULL;
	CAddressTable *prevTable = NULL;
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	int i;
	int j;
	int k;
	char s[1024];
	char *t;
	char *u;
	int ch = 0;
	char heading[30][32];
	char entry[30][256];
	CFont fontBook;
	CFont fontList;
	CFont fontLabel;
	CFont fontEnvelope;
	BOOL done = FALSE;
	BOOL openQuote = FALSE;
	BOOL isCity;
	BOOL isState;

	rewind(fp);

	// Read in single line from file
	memset(s, 0, 1024);
	for (i = 0; (i < 1023) && ((ch = fgetc(fp)) != EOF)
		&& (ch != '\n'); i++)
	{
		s[i] = (char)ch;
		if (s[i] == '\t')
		{
			s[i] = ' ';
		}
		//TRACE2("%d: %c\n", ch, s[i]);
	}
	s[i] = (char)ch;
	//TRACE2("%d: %c\n", ch, s[i]);
	if (s[i] == '\n')
	{
		s[i] = '\0';
	}

	if (ch == EOF)
	{
		fclose(fp);
		return FALSE;
	}

	k = 0;
	u = s;
	for (t = s; *t != '\0' && k < 29; t++)
	{
		if (*t == ',')
		{
			*t = '\0';
			sprintf(heading[k], "%s", u);
			//TRACE2("%d: %s\n", k, heading[k]);
			u = t + 1;
			k++;
		}
	}
	sprintf(heading[k], "%s", u);
	//TRACE2("%d: %s\n", k, heading[k]);
	k++;

	if (k > 30) // error: too many headings
	{
		fclose(fp);
		return FALSE;
	}

	// read entries into table
	i = 0;
	j = 0;
	while (TRUE)
	{
		ch = fgetc(fp);

		//TRACE2("%d: %c\n", ch, char(ch));
		entry[i][MIN(j, 255)] = (char)ch;

		if (ch == '"' && openQuote)
		{
			openQuote = FALSE;	
		}
		else if (ch == '"' && !openQuote)
		{
			openQuote = TRUE;
		}

		if ((ch == ',' && !openQuote) || (ch == '\n' && !openQuote) || ch == EOF)
		{
			entry[i][j] = '\0';
			rpRemoveQuotes(entry[i]);
			//TRACE2("(%d) %s\n", i, entry[i]);
			j = 0;
			i++;
		}
		else
		{
			j++;
		}

		if (i == k && ch != '\n' && ch != EOF || j >= 256)
		// illegal file format
		{
			fclose(fp);
			return FALSE;
		}

		if (i == k)
		// create a new table entry
		{
			ptrTable = new CAddressTable;

			for (i = 0; i < k && strcmp(heading[i], "First Name") != 0; i++);
			for (j = 0; j < k && strcmp(heading[j], "Last Name") != 0; j++);
			if (i < k && entry[i][0] != '\0' && j < k && entry[j][0] != '\0')
			// First Name, Last Name
			{
				sprintf(ptrTable->m_name, "%s %s", entry[i], entry[j]);
			}
			else if (i < k && entry[i][0] != '\0')
			// First Name
			{
				sprintf(ptrTable->m_name, "%s", entry[i]);
			}
			else if (j < k && entry[j][0] != '\0')
			// Last Name
			{
				sprintf(ptrTable->m_name, "%s", entry[j]);
			}
			//rpRemoveQuotes(ptrTable->m_name);

			for (i = 0; i < k && strcmp(heading[i], "Name") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Name
			{
				sprintf(ptrTable->m_name, "%s", entry[i]);
				//rpRemoveQuotes(ptrTable->m_name);
			}

			ptrTable->m_numberOfLinesBook++;
			ptrTable->m_numberOfLinesLabel++;

			for (i = 0; i < k && strcmp(heading[i], "E-mail Address") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// E-mail Address
			{
				sprintf(ptrTable->m_email[0], "%s", entry[i]);
				ptrTable->m_numberOfLinesBook++;
				for (i = 0; i < k && strcmp(heading[i], "Nickname") != 0; i++);
				if (i < k && entry[i][0] != '\0')
				// Nickname
				{
					sprintf(ptrTable->m_alias[0], "%s", entry[i]);
				}
			}

			for (i = 0; i < k && strcmp(heading[i], "Home Phone") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home Phone
			{
				sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				ptrTable->m_numberOfLinesBook++;
			}

			for (i = 0; i < k && strcmp(heading[i], "Mobile Phone") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Mobile Phone
			{
				if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;

			}

			for (i = 0; i < k && strcmp(heading[i], "Business Phone") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business Phone
			{
				if (ptrTable->m_phone[1][0] != '\0')
				{
					sprintf(ptrTable->m_phone[2], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
			}

			for (i = 0; i < k && strcmp(heading[i], "Business Fax") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business Fax
			{
				if (ptrTable->m_phone[2][0] != '\0')
				{
					sprintf(ptrTable->m_phone[3], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[1][0] != '\0')
				{
					sprintf(ptrTable->m_phone[2], "%s", entry[i]);
				}
				else if (ptrTable->m_phone[0][0] != '\0')
				{
					sprintf(ptrTable->m_phone[1], "%s", entry[i]);
				}
				else
				{
					sprintf(ptrTable->m_phone[0], "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
			}
			s[0] = '\0';
			for (i = 0; i < k && strcmp(heading[i], "Home Street") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home Street
			{
				//if (entry[i][0] == '"')
				//{
				//	t = entry[i] + 1;
				//}
				//else
				//{
				//	t = entry[i];
				//}
				//s[0] = *t;
				//j = 1;
				//for (t = t + 1; *t != '\0'; t++)
				j = 0;
				for (t = entry[i]; *t != '\0'; t++)
				{
					if (*t == '\n')
					{
						s[j] = ';';
						j++;
						s[j] = ' ';
						j++;
						while (*(t + 1) == ' ' || *(t + 1) == '\n')
						{
							t++;
						}
						ptrTable->m_numberOfLinesBook++;
						ptrTable->m_numberOfLinesLabel++;
					}
					//else if (*t == '"')
					//{
					//	s[j] = '\0';
					//	*t = '\0';
					//}
					else
					{
						s[j] = *t;
						j++;
					}
				}
				s[j] = '\0';
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
				sprintf(ptrTable->m_address[0], "%s", s);	
			}

			isCity = FALSE;
			for (i = 0; i < k && strcmp(heading[i], "Home City") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home City
			{
				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
				isCity = TRUE;
				sprintf(ptrTable->m_address[0], "%s", s);
			}

			isState = FALSE;
			for (i = 0; i < k && strcmp(heading[i], "Home State") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home State
			{
				//if (s[0] != '\0')
				//{
				//	sprintf(s, "%s %s", s, entry[i]);
				//}
				//else
				//{
				//	sprintf(s, "%s", entry[i]);
				//}
				//sprintf(ptrTable->m_address[0], "%s", s);

				if (s[0] != '\0' && isCity) //entry[i - 1] != '\0')
				{
					sprintf(s, "%s, %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city field is empty
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				isState = TRUE;
				sprintf(ptrTable->m_address[0], "%s", s);
			}

			for (i = 0; i < k && strcmp(heading[i], "Home Postal Code") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home Postal Code
			{
				//if (s[0] != '\0')
				//{
				//	sprintf(s, "%s  %s", s, entry[i]);
				//}
				//else
				//{
				//	sprintf(s, "%s", entry[i]);
				//}
				//sprintf(ptrTable->m_address[0], "%s", s);
				
				if (s[0] != '\0' && (isCity || isState)) //(entry[i - 2][0] != '\0' || entry[i - 1][0] != '\0'))
				{
					sprintf(s, "%s  %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city and state fields are empty but address field is not
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					ptrTable->m_numberOfLinesLabel++;
				}
				sprintf(ptrTable->m_address[0], "%s", s);	
			}

			for (i = 0; i < k && strcmp(heading[i], "Home Country/Region") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Home Country/Region
			{
				//sprintf(s, "%s; %s", s, entry[i]);
				//sprintf(ptrTable->m_address[0], "%s", s);
				//ptrTable->m_numberOfLinesBook++;
				//ptrTable->m_numberOfLinesLabel++;

				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				sprintf(ptrTable->m_address[0], "%s", s);
				ptrTable->m_numberOfLinesBook++;
				ptrTable->m_numberOfLinesLabel++;
			}
			//rpRemoveQuotes(ptrTable->m_address[0]);

			s[0] = '\0';
			for (i = 0; i < k && strcmp(heading[i], "Business Street") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business Street
			{
				j = 0;
				for (t = entry[i]; *t != '\0'; t++)
				{
					if (*t == '\n')
					{
						s[j] = ';';
						j++;
						s[j] = ' ';
						j++;
						while (*(t + 1) == ' ' || *(t + 1) == '\n')
						{
							t++;
						}
						ptrTable->m_numberOfLinesBook++;
						if (ptrTable->m_address[0][0] == '\0')
						{
							ptrTable->m_numberOfLinesLabel++;
						}
					}
					else
					{
						s[j] = *t;
						j++;
					}
				}
				s[j] = '\0';
				ptrTable->m_numberOfLinesBook++;
				if (ptrTable->m_address[0][0] == '\0')
				{
					ptrTable->m_numberOfLinesLabel++;
				}
				sprintf(ptrTable->m_address[1], "%s", s);	
			}

			isCity = FALSE;
			for (i = 0; i < k && strcmp(heading[i], "Business City") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business City
			{
				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				ptrTable->m_numberOfLinesBook++;
				if (ptrTable->m_address[0][0] == '\0')
				{
					ptrTable->m_numberOfLinesLabel++;
				}
				isCity = TRUE;
				sprintf(ptrTable->m_address[1], "%s", s);
			}

			isState = FALSE;
			for (i = 0; i < k && strcmp(heading[i], "Business State") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business State
			{
				//if (s[0] != '\0')
				//{
				//	sprintf(s, "%s %s", s, entry[i]);
				//}
				//else
				//{
				//	sprintf(s, "%s", entry[i]);
				//}
				//sprintf(ptrTable->m_address[1], "%s", s);
				

				if (s[0] != '\0' && isCity) //entry[i - 1] != '\0')
				{
					sprintf(s, "%s, %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city field is empty
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					if (ptrTable->m_address[0][0] == '\0')
					{
						ptrTable->m_numberOfLinesLabel++;
					}
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					if (ptrTable->m_address[0][0] == '\0')
					{
						ptrTable->m_numberOfLinesLabel++;
					}
				}
				isState = TRUE;
				sprintf(ptrTable->m_address[1], "%s", s);
			}

			for (i = 0; i < k && strcmp(heading[i], "Business Postal Code") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business Postal Code
			{
				//if (s[0] != '\0')
				//{
				//	sprintf(s, "%s  %s", s, entry[i]);
				//}
				//else
				//{
				//	sprintf(s, "%s", entry[i]);
				//}
				//sprintf(ptrTable->m_address[1], "%s", s);
				
				if (s[0] != '\0' && (isCity || isState)) //(entry[i - 2][0] != '\0' || entry[i - 1][0] != '\0'))
				{
					sprintf(s, "%s  %s", s, entry[i]);
				}
				else if (s[0] != '\0') // city and state fields are empty but address field is not
				{
					sprintf(s, "%s; %s", s, entry[i]);
					ptrTable->m_numberOfLinesBook++;
					if (ptrTable->m_address[0][0] == '\0')
					{
						ptrTable->m_numberOfLinesLabel++;
					}
				}
				else
				{
					sprintf(s, "%s", entry[i]);
					ptrTable->m_numberOfLinesBook++;
					if (ptrTable->m_address[0][0] == '\0')
					{
						ptrTable->m_numberOfLinesLabel++;
					}
				}
				sprintf(ptrTable->m_address[1], "%s", s);
			}

			for (i = 0; i < k && strcmp(heading[i], "Business Country/Region") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Business Country/Region
			{
				//sprintf(s, "%s; %s", s, entry[i]);
				//sprintf(ptrTable->m_address[1], "%s", s);
				//ptrTable->m_numberOfLinesBook++;
				//if (ptrTable->m_address[0][0] == '\0')
				//{
				//	ptrTable->m_numberOfLinesLabel++;
				//}
				if (s[0] != '\0')
				{
					sprintf(s, "%s; %s", s, entry[i]);
				}
				else
				{
					sprintf(s, "%s", entry[i]);
				}
				sprintf(ptrTable->m_address[1], "%s", s);
				ptrTable->m_numberOfLinesBook++;
				if (ptrTable->m_address[0][0] == '\0')
				{
					ptrTable->m_numberOfLinesLabel++;
				}
			}
			//rpRemoveQuotes(ptrTable->m_address[1]);
			
			s[0] = '\0';
			for (i = 0; i < k && strcmp(heading[i], "Notes") != 0; i++);
			if (i < k && entry[i][0] != '\0')
			// Notes
			{
			if (entry[i][0] == '"')
				{
					t = entry[i] + 1;
				}
				else
				{
					t = entry[i];
				}
				s[0] = *t;
				i = 0;
				j = 1;
				for (t = t + 1; *t != '\0' && i < 9; t++)
				{
					if (*t == '\n')
					{
						s[j] = '\0';
						j = 0;
						while (*(t + 1) == ' ')
						{
							t++;
						}
						if (i < 8)
						{
							sprintf(ptrTable->m_note[i], "%s", s);
							ptrTable->m_numberOfLinesBook++;
						}
						i++;
					}
					else if (*t == '"')
					{
						s[j] = '\0';
						*t = '\0';
					}
					else
					{
						s[j] = *t;
						j++;
					}
				}
				s[j] = '\0';
				ptrTable->m_numberOfLinesBook++;
				sprintf(ptrTable->m_note[i], "%s", s);
				
			}

			// move business address to first address position if no home address exists
			if (ptrTable->m_address[1][0] != '\0' && ptrTable->m_address[0][0] == '\0')
			{
				sprintf(ptrTable->m_address[0], "%s", ptrTable->m_address[1]);
				ptrTable->m_address[1][0] = '\0';
			}
			ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
			i = 0;
			j = 0;
			openQuote = FALSE;

		}

		if (ch == EOF)
		{
			break;
		}		
	}

	//First Name,Last Name,Middle Name,Name,Nickname,E-mail Address,Home Street,Home City,Home Postal Code,Home State,Home Country/Region,Home Phone,Home Fax,Mobile Phone,Personal Web Page,Business Street,Business City,Business Postal Code,Business State,Business Country/Region,Business Web Page,Business Phone,Business Fax,Pager,Company,Job Title,Department,Office Location,Notes
	rpUpdateAddressList();
	return TRUE;
}


void CAddressDoc::rpRemoveQuotes(char s[256])
// remove the first and last quote in the string s
// s might not contain any quotes
// the first quote must be in s[0] and the last quote must be in the space before '\0'
{
	int i = 0;
	char *t;
	BOOL comma = FALSE;
	BOOL newLine = FALSE;
	BOOL firstQuote = FALSE;
	BOOL lastQuote = FALSE;

	// remove first quote
	for (t = s; i < 256 && *t != '\0'; t++)
	{
		if (*t == '"' && t == s)
		{
			firstQuote = TRUE;
		}
		if (*t == ',' && firstQuote)
		{
			comma = TRUE;
		}
		if (*t == '\n' && firstQuote)
		{
			newLine = TRUE;
		}
		if (*t == '"' && *(t + 1) == '\0' && firstQuote && (comma || newLine))
		{
			lastQuote = TRUE;
			*t = '\0';
		}
	}

	if (firstQuote && (comma || newLine) && lastQuote)
	{
		sprintf(s, "%s", &s[1]);
	}

	return;
}

void CAddressDoc::rpSortName(void)
// takes all entries and re-inserts them into m_ptrTableList
// based on status of m_lastNameFirst
// group list is also reordered
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressTable *ptrTable = NULL;
	CAddressTable *curTable = NULL;
	CAddressTable *prevTable = NULL;
	CAddressTable *ptrDuplicateList;
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	CAddressGroup *ptrDuplicateGroupList;


	ptrDuplicateList = m_ptrTableList;
	m_ptrTableList = NULL;

	for (curTable = ptrDuplicateList; curTable != NULL;)
	{
		ptrTable = curTable;
		curTable = ptrTable->m_next;
		ptrTable->m_prev = NULL;
		ptrTable->m_next = NULL;
		//rpSortName(ptrTable);
		ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	rpUpdateAddressList();
	rpComputeDrawPositions();

	// re-order group list
	ptrDuplicateGroupList = m_ptrGroupList;
	m_ptrGroupList = NULL;

	for (curGroup = ptrDuplicateGroupList; curGroup != NULL;)
	{
		ptrGroup = curGroup;
		curGroup = ptrGroup->m_next;
		ptrGroup->m_prev = NULL;
		ptrGroup->m_next = NULL;
		ptrGroup->rpInsertGroup(&m_ptrGroupList, m_lastNameFirst);
	}

	// m_groupSelected may need changing based on re-order
	for (ptrGroup = m_groupSelected;
		ptrGroup != NULL && ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_groupName, ptrGroup->m_prev->m_groupName) == 0;
		ptrGroup = ptrGroup->m_prev);
	m_groupSelected = ptrGroup;

	return;
}


/**
void CAddressDoc::rpSortName(CAddressTable *ptrTable)
// arranges s based on selection of m_lastNameFirst
{
	int numCommas;
	int j;
	char *t;
	char *u;
	char firstName[256];
	char lastName[256];

	// count the commas -- do not count \\, sequence
	j = 0;
	for (t = s; *t != '\0'; t++)
	{
		if (*t == ',' && j >= 2 && *(t - 1) != '\\' && *(t - 2) != '\\'
			|| *t == ',' && j < 2)
		{
			numCommas++;
		}
	}

	if (numCommas <= 1)

	u = s;
	for (t = s; *t != '\0' && *t != ','; t++);
	if (*t == ',')
	{
		*t = '\0';
		t++;
		for (; *t == ' ' && *t != '\0'; t++);
		sprintf(firstName, "%s", t);
		sprintf(lastName, "%s", u);
	}
	else 
	{
		for (; *t != ' ' && t != s; t--);
		if (*t == ' ')
		{
			*t = '\0';
			t++;
			sprintf(firstName, "%s", u);
			sprintf(lastName, "%s", t);
		}
		else  // one name entry
		{
			sprintf(firstName, "%s", u);
			lastName[0] = '\0';
		}
	}

	// make certain there are no commas in firstName
	// if there are commas in first name designate by char series \\,
	sprintf(s, "%s", firstName);
	j = 0;
	for (t = s; *t != '\0'; t++)
	{
		if (*t == ',')
		{
			firstName[j] = '\\';
			firstName[j + 1] = '\\';
			firstName[j + 2] = ',';
			j += 3;
		}
		else
		{
			firstName[j] = *t;
			j++;
		}
	}
	firstName[j] = '\0';

	if (m_lastNameFirst == TRUE && lastName[0] != '\0')
	{
		sprintf(s, "%s, %s", lastName, firstName);
	}
	else if (m_lastNameFirst == FALSE && lastName[0] != '\0')
	{
		sprintf(s, "%s %s", firstName, lastName);
	}
	else
	{
		sprintf(s, "%s", firstName);
	}

	return;
}
**/


void CAddressDoc::rpFreeDoc(void)
// Free memory for all CAddressTable nodes.
{
	CAddressTable *prevTable = NULL;
	CAddressTable *curTable;
	CAddressAlias *prevAlias = NULL;
	CAddressAlias *curAlias;
	CAddressGroup *prevGroup = NULL;
	CAddressGroup *curGroup;

	for (curTable = m_ptrTableList; curTable != NULL;)
	{
		prevTable = curTable;
		curTable = curTable->m_next;
		delete prevTable;
	}

	for (curAlias = m_ptrAliasList; curAlias != NULL;)
	{
		prevAlias = curAlias;
		curAlias = curAlias->m_next;
		delete prevAlias;
	}

	prevGroup = NULL;
	for (curGroup = m_ptrGroupList; curGroup != NULL;)
	{
		prevGroup = curGroup;
		curGroup = curGroup->m_next;
		delete prevGroup;
	}

	m_ptrTableList = NULL;
	m_ptrAliasList = NULL;
	m_ptrGroupList = NULL;
	m_groupSelected = NULL;
	m_returnAddress[0] = '\0';
	m_firstLabel = 0;
	m_lastLabel = 0;

	rpComputeDrawPositions();
	rpUpdateAddressList();

	return;
}


CSize CAddressDoc::rpGetDocSize(void)
{

	return m_sizeDoc;
	/***
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}

	CAddressTable *ptrTable;
	int i;
	CPrintDialog *pPrintDlg;
	//int j;


	i = 0;	
	for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		i += ptrTable->m_numberOfLinesBook;
	}

	m_tableSize[1] = m_lineSpacing * i;


	m_sizeDoc = CSize(m_tableSize[0] + m_offsetBorder * 2, m_tableSize[1] + m_offsetBorder * 2);
	return m_sizeDoc;
	**/

}

BOOL CAddressDoc::OnOpenDocument(LPCTSTR lpszPathName) 
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}
	FILE *fp;
	char s[256];
	CWnd *pCBox;

	if (!CDocument::OnOpenDocument(lpszPathName))
	{
		return FALSE;
	}
	
	// TODO: Add your specialized creation code here
	if ((fp = fopen((LPCTSTR)lpszPathName, "r")) == NULL)
	{
		sprintf(s, "File open failed.  Unable to find file:\n   %s", (LPCTSTR)lpszPathName);
		AfxMessageBox(s);
		return FALSE;
	}
	else
	{
		rpFreeDoc();
		if (rpReadDoc(fp) == FALSE)
		{
			sprintf(s, "File open failed due to incorrect file format.");
			AfxMessageBox(s);
			fclose(fp);
			pMainFrame->SendMessage(WM_COMMAND, ID_FILE_NEW, (LPARAM)0);
			return FALSE;
		}
		fclose(fp);
	}

	// after file is read, sort names in m_ptrTableList based on value of m_lastNameFirst
	rpSortName();

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", m_lineSpacing);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();

		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
		sprintf(s, "%d", m_columnWidth);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", m_lineSpacingList);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", m_lineSpacingEnvelope);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
	m_return = FALSE;
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		if (m_returnAddress[0] == '\0')
		{
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}
		else
		{
			pCBox->ShowWindow(SW_SHOW);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}

	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		if (m_returnAddress[0] == '\0')
		{
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}
		else
		{
			pCBox->ShowWindow(SW_SHOW);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	rpComputeDrawPositions();

	//pMainFrame->rpUpdateGroupMenu();
	rpUpdateGroupMenu();

	return TRUE;
}

BOOL CAddressDoc::OnSaveDocument(LPCTSTR lpszPathName) 
{
	//sprintf(m_fileName, "%s", (LPCTSTR)lpszPathName);
	rpWriteDoc(lpszPathName);
	SetModifiedFlag(FALSE);
	return TRUE;

}

void CAddressDoc::OnWriteUnixMailrc()
{
	FILE *fp;
	CAddressAlias *ptrAlias;
	CAddressAlias *curAlias;
	CAddressAlias *prevAlias;
	CAddressGroup *ptrGroup;
	CAddressAlias *ptrGroupAliasList = NULL;
	CAddressTable *ptrTable;
	char s[256];
	int j;

	// TODO: Add your command handler code here
	CFileDialog writeUnixMailrc
		(FALSE,    // save
		NULL,    // no default extension
		NULL,    // file name
		OFN_HIDEREADONLY,
		"All files (*.*)|*.*||");

	writeUnixMailrc.m_ofn.lpstrTitle = "Save in Unix .mailrc format";

	if (writeUnixMailrc.DoModal() == IDOK)
	{
		// create alias list
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			for (j = 0; j < 9; j++)
			{
				if (ptrTable->m_alias[j][0] != '\0')
				{
					if (ptrTable->m_email[j][0] != '\0')
					{
						ptrAlias = new CAddressAlias;
						sprintf(ptrAlias->m_alias, "%s", ptrTable->m_alias[j]);
						sprintf(ptrAlias->m_email, "%s", ptrTable->m_email[j]);
						ptrAlias->rpInsertAlias(&m_ptrAliasList);
					}
				}
			}

		}

		// create group list based on emails or aliases
		for (ptrGroup = m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
		{
			for (j = 0; j < 9; j++)
			{
				if (ptrGroup->m_ptrTable->m_email[j][0] != '\0')
				{
					ptrAlias = new CAddressAlias;
					sprintf(ptrAlias->m_alias, "%s", ptrGroup->m_groupName);
					if (ptrGroup->m_ptrTable->m_alias[j][0] != '\0')
					{
						sprintf(ptrAlias->m_email, "%s", ptrGroup->m_ptrTable->m_alias[j]);
					}
					else
					{
						sprintf(ptrAlias->m_email, "%s", ptrGroup->m_ptrTable->m_email[j]);
					}
					ptrAlias->rpInsertAlias(&ptrGroupAliasList);
				}
			}

		}


		fp = fopen((LPCSTR)writeUnixMailrc.GetPathName(), "w");
		if (fp == NULL)
		{
			return;
		}

		fprintf(fp, "# .mailrc\n");
		fprintf(fp, "# Set some default behavior for mail.\n");
		fprintf(fp, "#\n");
		fprintf(fp, "# See 'man mail' for more information.\n");
		fprintf(fp, "#\n");
		fprintf(fp, "# Set the following behaviors:\n");
		fprintf(fp, "#    ask        - make mail ask for a subject.\n");
		fprintf(fp, "#    askcc      - prompt user for carbon copy recipients.\n");
		fprintf(fp, "#    crt=23     - messages longer than 23 lines need a pager.\n");
		fprintf(fp, "#    hold       - store messages in the system spool directory.\n");
		fprintf(fp, "#\n");
		fprintf(fp, "set ask askcc crt=23 hold\n");
		fprintf(fp, "#");

		s[0] = '\0';
		for (ptrAlias = m_ptrAliasList; ptrAlias != NULL; ptrAlias = ptrAlias->m_next)
		{
			if (strcmp(ptrAlias->m_alias, s) == 0)
			{
				fprintf(fp, " %s", ptrAlias->m_email);
			}
			else
			{
				fprintf(fp, "\nalias %s %s", ptrAlias->m_alias, ptrAlias->m_email);
			}
			sprintf(s, ptrAlias->m_alias);
		}
		fprintf(fp, "\n");

		s[0] = '\0';
		for (ptrAlias = ptrGroupAliasList; ptrAlias != NULL; ptrAlias = ptrAlias->m_next)
		{
			if (strcmp(ptrAlias->m_alias, s) == 0
				&& ptrAlias->m_next != NULL
				&& strcmp(ptrAlias->m_alias, ptrAlias->m_next->m_alias) == 0)
			{
				fprintf(fp, "\t%s \\\n", ptrAlias->m_email);
			}
			else if (strcmp(ptrAlias->m_alias, s) == 0)
			{
				fprintf(fp, "\t%s\n", ptrAlias->m_email);
			}
			else
			{
				fprintf(fp, "#\nalias %s \\\n\t%s \\\n", ptrAlias->m_alias, ptrAlias->m_email);
			}			
			sprintf(s, ptrAlias->m_alias);
		}

		//for (ptrAlias = m_ptrGroupList; ptrAlias != NULL; ptrAlias = ptrAlias->m_next)
		//{
		//	if (strcmp(ptrAlias->m_alias, s) == 0
		//		&& ptrAlias->m_next != NULL
		//		&& strcmp(ptrAlias->m_alias, ptrAlias->m_next->m_alias) == 0)
		//	{
		//		fprintf(fp, "\t%s \\\n", ptrAlias->m_email);
		//	}
		//	else if (strcmp(ptrAlias->m_alias, s) == 0)
		//	{
		//		fprintf(fp, "\t%s\n", ptrAlias->m_email);
		//	}
		//	else
		//	{
		//		fprintf(fp, "#\nalias %s \\\n\t%s \\\n", ptrAlias->m_alias, ptrAlias->m_email);
		//	}
		//	sprintf(s, ptrAlias->m_alias);
		//}

		fclose(fp);

		// remove alias list
		prevAlias = NULL;
		for (curAlias = m_ptrAliasList; curAlias != NULL;)
		{
			prevAlias = curAlias;
			curAlias = curAlias->m_next;
			delete prevAlias;
		}
		m_ptrAliasList = NULL;

		// remove group alias list
		prevAlias = NULL;
		for (curAlias = ptrGroupAliasList; curAlias != NULL;)
		{
			prevAlias = curAlias;
			curAlias = curAlias->m_next;
			delete prevAlias;
		}

	}
	return;
}

void CAddressDoc::OnUpdateFilePrint(CCmdUI* pCmdUI)
// Printing commands are disabled
{
	// TODO: Add your command update UI handler code here
	//pCmdUI->Enable(FALSE);
}

void CAddressDoc::OnUpdateFilePrintPreview(CCmdUI* pCmdUI) 
// Printing commands are disabled
{
	// TODO: Add your command update UI handler code here
	//pCmdUI->Enable(FALSE);
}

void CAddressDoc::OnUpdateFilePrintSetup(CCmdUI* pCmdUI) 
// Printing commands are disabled
{
	// TODO: Add your command update UI handler code here
	//pCmdUI->Enable(FALSE);
}


BOOL CAddressDoc::rpUpdateGroupMenu()
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CMenu *menu;
	char groupName[30][64];
	int k = 0;
	int i;
	static int prevGroupCount = 30;
	CAddressGroup *ptrGroup;

	for (ptrGroup = m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			pMainFrame->m_allGroupMembersSelected[k] = TRUE;
			sprintf(groupName[k], "%s", ptrGroup->m_groupName);
			k++;
		}
		if (pMainFrame->m_allGroupMembersSelected[k - 1] == TRUE && ptrGroup->m_ptrTable->m_selected == FALSE)
		{
			pMainFrame->m_allGroupMembersSelected[k - 1] = FALSE;
		}
	}

	if ((menu = pMainFrame->GetMenu()) == NULL)
	{
		return FALSE;
	}

	TRACE1("Menu %d\n", menu);

	menu->DeleteMenu(ID_GROUP_ONE, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_TWO, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_THREE, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_FOUR, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_FIVE, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_SIX, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_SEVEN, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_EIGHT, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_NINE, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_TEN, MF_BYCOMMAND);
	menu->DeleteMenu(ID_GROUP_ELEVEN, MF_BYCOMMAND);

	if (k == 0)
	{
		menu->ModifyMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_TWELVE, "(empty)");
	}
	else
	{
		for (i = 0; i < k - 1; i++)
		{
			switch (i)
			{
				case 0:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_ONE, groupName[i]);
					break;
				case 1:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_TWO, groupName[i]);
					break;
				case 2:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_THREE, groupName[i]);
					break;
				case 3:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_FOUR, groupName[i]);
					break;
				case 4:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_FIVE, groupName[i]);
					break;
				case 5:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_SIX, groupName[i]);
					break;
				case 6:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_SEVEN, groupName[i]);
					break;
				case 7:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_EIGHT, groupName[i]);
					break;
				case 8:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_NINE, groupName[i]);
					break;
				case 9:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_TEN, groupName[i]);
					break;
				case 10:
					menu->InsertMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_ELEVEN, groupName[i]);
					break;
				//case 11:
				//	break;
			}
		}
		menu->ModifyMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_TWELVE, groupName[k - 1]);
	}

	return TRUE;
}

int CAddressDoc::rpComputeDrawPositions()
// set position of each entry in draw window based on desired output
// return number of pages in document
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CAddressTable *ptrTable;
	int linesPerPage;
	int columnsPerPage;
	int columnOffset;
	int pages;
	int offsetX;
	int offsetY;
	int i;
	int j;
	int k;
	int m;
	int n;
	int p;
	int charsPerLine;
	int numberOfLines;
	char s[4096];
	char *u;
	char *v;
	char *w;
	int charName;
	int charAddress[10];
	int charEmail[10];
	int charPhone[10];
	int charNote[10];
	BOOL enableDelete = FALSE;
	BOOL checkSelectAll = TRUE;
	BOOL checkSelectNone = TRUE;
	CDC dc;
	CPrintDialog dlg(FALSE);
	CWnd *pCBox;
	CWinApp* app = AfxGetApp();
	CRect curCellRect;
	CFont *ptrOldFont;
	CDC *pDC;
	pDC = pView->GetDC( );

	app->GetPrinterDeviceDefaults(&dlg.m_pd);

	LPDEVMODE lp = (LPDEVMODE) ::GlobalLock(dlg.m_pd.hDevMode);
	ASSERT(lp);

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		lp->dmPaperSize = m_paperSizeBook;
		lp->dmOrientation = m_paperOrientationBook;
	}
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		lp->dmPaperSize = m_paperSizeList;
		lp->dmOrientation = m_paperOrientationList;
	}
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		lp->dmPaperSize = m_paperSizeLabel;
		lp->dmOrientation = DMORIENT_PORTRAIT;
		//lp->dmOrientation = m_paperOrientationLabel;
		//lp->dmPaperSize = DMPAPER_LETTER;
	}
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		lp->dmPaperSize = m_paperSizeEnvelope;
		lp->dmOrientation = m_paperOrientationEnvelope;
	}

	::GlobalUnlock(dlg.m_pd.hDevMode);

	dlg.CreatePrinterDC();
	dc.Attach(dlg.m_pd.hDC);

	m_pageWidth = dc.GetDeviceCaps(PHYSICALWIDTH) * 100 / dc.GetDeviceCaps(LOGPIXELSX);
	m_pageHeight = dc.GetDeviceCaps(PHYSICALHEIGHT) * 100 / dc.GetDeviceCaps(LOGPIXELSY);

	m_printableWidth = dc.GetDeviceCaps(HORZRES) * 100 / dc.GetDeviceCaps(LOGPIXELSX);
	m_printableHeight = dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY);

	m_physicalOffsetX = dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX);
	m_physicalOffsetY = dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY);


	m_margin[0] = MAX(m_margin[0], m_physicalOffsetX);
	m_margin[1] = MAX(m_margin[1], m_physicalOffsetY);
	m_margin[2] = MAX(m_margin[2], m_pageWidth - m_physicalOffsetX - m_printableWidth);
	m_margin[3] = MAX(m_margin[3], m_pageHeight - m_physicalOffsetY - m_printableHeight);


	TRACE2("rpComputeDrawPositions():  m_printableWidth = %d   m_printableHeight = %d\n",
		m_printableWidth, m_printableHeight);
	TRACE2("m_pageWidth = %d   m_pageHeight = %d\n",
		m_pageWidth, m_pageHeight);
	TRACE2("m_physicalOffsetX = %d   m_physicalOffsetY = %d\n",
		m_physicalOffsetX, m_physicalOffsetY);

	// first select all names in left dialog box.  Some may be unselected later
	pCBox = (CComboBox *)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
	i = ((CListBox *)pCBox)->GetCount() - 1;
	if (i > 0)
	{
		((CListBox *)pCBox)->SelItemRange(TRUE, 0, i);
	}
	else
	{
		((CListBox *)pCBox)->SetSel(0, TRUE);
	}

	// draw address book
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		linesPerPage = (m_pageHeight - m_margin[1] - m_margin[3]) / m_lineSpacing;
		columnsPerPage = MAX(1, (m_pageWidth - m_margin[0] - m_margin[2]) / m_columnWidth);
		columnOffset = (m_pageWidth - m_margin[0] - m_margin[2]) / columnsPerPage;

		j = 0;
		i = 0;
		offsetX = m_margin[0] - m_physicalOffsetX;
		offsetY = m_margin[1] - m_physicalOffsetY;

		pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
		k = 0;
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			//if (((CListBox *)pCBox)->GetSel(k))
			if (ptrTable->m_selected)
			{
				//((CListBox *)pCBox)->SetSel(k, TRUE);
				enableDelete = TRUE;
				checkSelectNone = FALSE;
				i += ptrTable->m_numberOfLinesBook;
				if (i > linesPerPage)
				{
					j++;
					i = ptrTable->m_numberOfLinesBook;
					offsetY = (j / columnsPerPage) * m_printableHeight
						+ m_margin[1] - m_physicalOffsetY;
					offsetX = m_margin[0] - m_physicalOffsetX
						+ (j * columnOffset) % (columnsPerPage * columnOffset);
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += ptrTable->m_numberOfLinesBook * m_lineSpacing;			
				}
				else
				{
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += ptrTable->m_numberOfLinesBook * m_lineSpacing;
				}
				// set lower right corner of draw bounding box
				ptrTable->m_drawCoord[2] = offsetX + columnOffset;
				ptrTable->m_drawCoord[3] = -offsetY;
			}
			else
			{
				checkSelectAll = FALSE;
				((CListBox *)pCBox)->SetSel(k, FALSE);
				ptrTable->m_drawCoord[0] = 0;
				ptrTable->m_drawCoord[1] = 0;
				ptrTable->m_drawCoord[2] = 0;
				ptrTable->m_drawCoord[3] = 0;
			}
			k++;
		}
		pages = (j / columnsPerPage) + 1;
		m_sizeDoc = CSize(
			m_pageWidth,
			pages * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	}


	// draw address list
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		ptrOldFont = pDC->SelectObject(&m_fontList);
		linesPerPage = (m_pageHeight - m_margin[1] - m_margin[3]) / m_lineSpacingList;

		j = 0;
		i = 0;
		offsetX = m_margin[0] - m_physicalOffsetX;
		offsetY = m_margin[1] - m_physicalOffsetY;

		pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
		k = 0;
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			//if (((CListBox *)pCBox)->GetSel(k))
			if (ptrTable->m_selected)
			{
				//((CListBox *)pCBox)->SetSel(k, TRUE);
				enableDelete = TRUE;
				checkSelectNone = FALSE;
				for (p = 0; p < 100; p++)
				{
					ptrTable->m_listLineBreak[p] = NULL;
				}

				if (m_lastNameFirst == FALSE)
				{
					sprintf(ptrTable->m_list, "%s", ptrTable->m_name);
				}
				else
				{
					sprintf(ptrTable->m_list, "%s", ptrTable->m_nameLastFirst);
				}

				charName = strlen(ptrTable->m_name);
				for (m = 0; m < 9 && ptrTable->m_address[m][0] != '\0'; m++)
				{
					sprintf(ptrTable->m_list, "%s * %s", ptrTable->m_list, ptrTable->m_address[m]);
					charAddress[m] = strlen(ptrTable->m_address[m]);
					ptrTable->m_address[m][255] = '\0';
				}
				charAddress[m] = -1;
				for (m = 0; m < 9 && ptrTable->m_email[m][0] != '\0'; m++)
				{
					sprintf(ptrTable->m_list, "%s * %s", ptrTable->m_list, ptrTable->m_email[m]);
					charEmail[m] = strlen(ptrTable->m_email[m]);
					ptrTable->m_email[m][255] = '\0';
				}
				charEmail[m] = -1;
				for (m = 0; m < 9 && ptrTable->m_phone[m][0] != '\0'; m++)
				{
					sprintf(ptrTable->m_list, "%s * %s", ptrTable->m_list, ptrTable->m_phone[m]);
					charPhone[m] = strlen(ptrTable->m_phone[m]);
					ptrTable->m_phone[m][255] = '\0';
				}
				charPhone[m] = -1;
				
				v = ptrTable->m_list;
				v = v + strlen(ptrTable->m_list);
				for (m = 0; m < 9 && ptrTable->m_note[m][0] != '\0'; m++)
				{
					sprintf(ptrTable->m_list, "%s * %s", ptrTable->m_list, ptrTable->m_note[m]);
					charNote[m] = strlen(ptrTable->m_note[m]);
					ptrTable->m_note[m][255] = '\0';
				}
				charNote[m] = -1;
				
				curCellRect.left = 0;
				curCellRect.top = 0;
				curCellRect.right = m_printableWidth;
				curCellRect.bottom = curCellRect.top - m_lineSpacingList * 2;
				pDC->DrawText(ptrTable->m_list, curCellRect, DT_CALCRECT);
				numberOfLines = 1 + (curCellRect.right / (m_pageWidth - m_margin[0] - m_margin[2]));
				ptrTable->m_listLineBreak[0] = ptrTable->m_list;
				

				if (numberOfLines > 1)
				{
					charsPerLine = strlen(ptrTable->m_list) * (m_pageWidth - m_margin[0] - m_margin[2]) / curCellRect.right;
					TRACE3("%s: strlen = %d   charsPerLine = %d   ",
						ptrTable->m_name, strlen(ptrTable->m_list), charsPerLine);
					TRACE2("numberOfLines = %d   curCellRect.right = %d\n", numberOfLines, curCellRect.right);
					numberOfLines = 1; // recalculate number of lines with different spacing

					n = 0;
					p =	1;
					for (u = ptrTable->m_list; *u != '\0'; u++)
					{
						n++;
						if (n == charsPerLine)
						{
							if (p == 1)
							{
								charsPerLine -= 5;  // subtract 5 spaces for line indentation
							}
							w = u;
							if (u > v) // into note part of list
							{
								// first try to find * and see if remaining text can be placed on next line
								for (; *u != '*' && n != 0; u--)
								{
									n--;
								}
								if (strlen(u) > (UINT)charsPerLine + 2)
								{
									// try to break at space
									u = w;
									n = charsPerLine;
									for (; *u != ' ' && n != 0; u--)
									{
										n--;
									}
									if (n != 0 && *(u - 1) == '*')
									// note without spaces to right margin
									// highly unlikely to occur
									{
										n = 0;  // break at end at end of line
									}
								}

							}
							else
							{
								for (; *u != '*' && *u != ';' && n != 0; u--)
								{
									n--;
								}
							}

							if (n == 0) // no spaces in note
							{
								u = w; // set break to char at end of line
							}
							
							for (u = u + 1; *u == ' '; u++);
							if (*u != '\0')
							{
								if (*(u - 1) == ' ')
								{
									*(u - 1) = '\0';
								}
								else
								// called only if break was set at end of line
								// and line contains no spaces--highly unlikely to occur
								{
									sprintf(s, "%s", u);
									sprintf(&ptrTable->m_list[u - ptrTable->m_list + 1], "%s", s);
									*u = '\0';
									u++;
								}
								ptrTable->m_listLineBreak[p] = u;
								TRACE1("   %s\n", u);
								numberOfLines++;
								p++;
								n = 0;  // indentation for next line
							}
							else
							{
								u--; // set to end of string for next iteration
							}
						}
					}
				}
				
				/**
				if (numberOfLines > 1)
				{
					charsPerLine = strlen(s) * (m_pageWidth - m_margin[0] - m_margin[2]) / curCellRect.right;
					n = charName;
					TRACE3("%s: strlen = %d   charsPerLine = %d   ",
						ptrTable->m_name, strlen(s), charsPerLine);
					TRACE2("numberOfLines = %d   curCellRect.right = %d\n", numberOfLines, curCellRect.right);
					numberOfLines = 1; // recalculate number of lines with different spacing
					for (m = 0; m < 10 && charAddress[m] != -1; m++)
					{
						n += charAddress[m] + 3;
						if (n > charsPerLine)
						{
							ptrTable->m_address[m][255] = '\n';
							n = charAddress[m] + 5; // 5 is indentation for next line
							numberOfLines++;
						}
					}
					for (m = 0; m < 10 && charEmail[m] != -1; m++)
					{
						n += charEmail[m] + 3;
						if (n > charsPerLine)
						{
							ptrTable->m_email[m][255] = '\n';
							n = charEmail[m] + 5; // 5 is indentation for next line
							numberOfLines++;
						}
					}
					for (m = 0; m < 10 && charPhone[m] != -1; m++)
					{
						n += charPhone[m] + 3;
						if (n > charsPerLine)
						{
							ptrTable->m_phone[m][255] = '\n';
							n = charPhone[m] + 5; // 5 is indentation for next line
							numberOfLines++;
						}
					}
					for (m = 0; m < 10 && charNote[m] != -1; m++)
					{
						n += charNote[m] + 3;
						if (n > charsPerLine)
						{
							ptrTable->m_note[m][255] = '\n';
							n = charNote[m] + 5; // 5 is indentation for next line
							numberOfLines++;
						}
					}
				}
				**/
				

				i += numberOfLines;
				//i++;

				if (i > linesPerPage)
				{
					TRACE3("%d: i = %d   linesPerPage = %d\n", j, i, linesPerPage);
					j++;
					//i = 1;
					i = numberOfLines;
					offsetY = j * m_printableHeight
						+ m_margin[1] - m_physicalOffsetY;
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += m_lineSpacingList * numberOfLines;			
				}
				else
				{
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += m_lineSpacingList * numberOfLines;
				}
				// set lower right corner of draw bounding box
				ptrTable->m_drawCoord[2] = offsetX + m_printableWidth;
				ptrTable->m_drawCoord[3] = -offsetY;
			}
			else
			{
				checkSelectAll = FALSE;
				((CListBox *)pCBox)->SetSel(k, FALSE);
				ptrTable->m_drawCoord[0] = 0;
				ptrTable->m_drawCoord[1] = 0;
				ptrTable->m_drawCoord[2] = 0;
				ptrTable->m_drawCoord[3] = 0;
			}
			k++;
		}
		pages = j + 1;
		m_sizeDoc = CSize(
			m_pageWidth,
			pages * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
		pDC->SelectObject(ptrOldFont);
	}

		// draw avery label
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		//linesPerPage = m_labelDimension[1];
		//columnsPerPage = m_labelDimension[0];
		//columnOffset = m_labelColumnWidth;

		j = m_firstLabel / m_labelDimension[1];   // column
		i = m_firstLabel % m_labelDimension[1];   // row

		// get position (upper left corner) of start label
		offsetY = ((j / m_labelDimension[0]) * m_printableHeight
			+ m_labelTopMargin - m_physicalOffsetY)
			+ ((m_firstLabel % m_labelDimension[1]) * m_labelHeight);
		offsetX = m_labelLeftMargin - m_physicalOffsetX
			+ (j * m_labelColumnWidth) % (m_labelDimension[0] * m_labelColumnWidth);

		pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
		k = 0;
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		// set position (upper left corner) of label for each selected entry in table
		{
			if (ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1)
			{
				enableDelete = TRUE;
				checkSelectNone = FALSE;
				i++;
				if (i > m_labelDimension[1])
				{
					j++;
					i = 1;
					offsetY = (j / m_labelDimension[0]) * m_printableHeight
						+ m_labelTopMargin - m_physicalOffsetY;
					offsetX = m_labelLeftMargin - m_physicalOffsetX
						+ (j * m_labelColumnWidth) % (m_labelDimension[0] * m_labelColumnWidth);
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += m_labelHeight;			
				}
				else
				{
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += m_labelHeight;
				}
				// set lower right corner of draw bounding box
				ptrTable->m_drawCoord[2] = offsetX + m_labelWidth;
				ptrTable->m_drawCoord[3] = -offsetY;
			}
			else
			{
				//if (!((CListBox *)pCBox)->GetSel(k) && ptrTable->m_numberOfLinesLabel > 1)
				if (!ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1) 
				{
					checkSelectAll = FALSE;
				}
				((CListBox *)pCBox)->SetSel(k, FALSE);
				ptrTable->m_drawCoord[0] = 0;
				ptrTable->m_drawCoord[1] = 0;
				ptrTable->m_drawCoord[2] = 0;
				ptrTable->m_drawCoord[3] = 0;
			}


			k++;
		}
		pages = (j / m_labelDimension[0]) + 1;
		m_sizeDoc = CSize(
			m_pageWidth,
			pages * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	}

	/***
	// draw avery label
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		linesPerPage = 60;
		columnsPerPage = 3;
		columnOffset = m_labelColumnWidth;

		j = m_firstLabel / 10;
		i = (m_firstLabel % 10) * 6;

		offsetY = ((j / columnsPerPage) * m_printableHeight
			+ m_labelTopMargin - m_physicalOffsetY)
			+ ((m_firstLabel % 10) * 100);
		offsetX = m_labelLeftMargin - m_physicalOffsetX
			+ (j * columnOffset) % (columnsPerPage * columnOffset);
		//offsetX = 50 - m_physicalOffsetX
		//	+ (j * columnOffset) % (columnsPerPage * columnOffset);
		//offsetX = 50 - m_physicalOffsetX;
		//offsetY = 50 - m_physicalOffsetY;

		pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
		k = 0;
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			//if (((CListBox *)pCBox)->GetSel(k) && ptrTable->m_numberOfLinesLabel > 1)
			if (ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1)
			{
				//((CListBox *)pCBox)->SetSel(k, TRUE);
				enableDelete = TRUE;
				checkSelectNone = FALSE;
				i += 6;
				if (i > linesPerPage)
				{
					j++;
					i = 6;
					offsetY = (j / columnsPerPage) * m_printableHeight
						+ m_labelTopMargin - m_physicalOffsetY;
					offsetX = m_labelLeftMargin - m_physicalOffsetX
						+ (j * columnOffset) % (columnsPerPage * columnOffset);
					//offsetX = 50 - m_physicalOffsetX
					//	+ (j * columnOffset) % (columnsPerPage * columnOffset);
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += 100;			
				}
				else
				{
					ptrTable->m_drawCoord[0] = offsetX;
					ptrTable->m_drawCoord[1] = -offsetY;
					offsetY += 100;
				}
			}
			else
			{
				//if (!((CListBox *)pCBox)->GetSel(k) && ptrTable->m_numberOfLinesLabel > 1)
				if (!ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1) 
				{
					checkSelectAll = FALSE;
				}
				((CListBox *)pCBox)->SetSel(k, FALSE);
			}


			k++;
		}
		pages = (j / columnsPerPage) + 1;
		m_sizeDoc = CSize(
			m_pageWidth,
			pages * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	}
	**/

	// draw envelope
	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		linesPerPage = m_pageHeight / m_lineSpacingEnvelope;

		j = 0;
		k = 0;
		offsetX = m_physicalOffsetX + (2 * m_printableWidth / 5);

		pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			//if (((CListBox *)pCBox)->GetSel(k) && ptrTable->m_numberOfLinesLabel > 1)
			if (ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1)
			{
				//((CListBox *)pCBox)->SetSel(k, TRUE);
				enableDelete = TRUE;
				checkSelectNone = FALSE;
				offsetY = j * m_printableHeight
					+ m_physicalOffsetY + (m_printableHeight / 2);
				ptrTable->m_drawCoord[0] = offsetX;
				ptrTable->m_drawCoord[1] = -offsetY;
				// set lower right corner of draw bounding box
				ptrTable->m_drawCoord[2] = m_physicalOffsetX + m_printableWidth;
				ptrTable->m_drawCoord[3] = -(j * m_printableHeight
					+ m_physicalOffsetY + m_printableHeight);
				j++;
			}
			else
			{
				//if (!((CListBox *)pCBox)->GetSel(k) && ptrTable->m_numberOfLinesLabel > 1)
				if (!ptrTable->m_selected && ptrTable->m_numberOfLinesLabel > 1)
				{
					checkSelectAll = FALSE;
				}
				((CListBox *)pCBox)->SetSel(k, FALSE);
				ptrTable->m_drawCoord[0] = 0;
				ptrTable->m_drawCoord[1] = 0;
				ptrTable->m_drawCoord[2] = 0;
				ptrTable->m_drawCoord[3] = 0;
			}
			k++;
		}
		pages = j;
		m_sizeDoc = CSize(
			m_pageWidth,
			pages * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	}

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_SELECT_ALL);
	if (m_ptrTableList == NULL)
	{
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(FALSE);
	}
	else
	{
		pCBox->EnableWindow(TRUE);
		((CButton *)pCBox)->SetCheck(checkSelectAll);
	}

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_SELECT_NONE);
	if (m_ptrTableList == NULL)
	{
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(FALSE);
	}
	else
	{
		pCBox->EnableWindow(TRUE);
		((CButton *)pCBox)->SetCheck(checkSelectNone);
	}

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_DELETE);
	pCBox->EnableWindow(enableDelete);

	pView->SetScrollSizes(MM_LOENGLISH, pView->GetDocument()->rpGetDocSize());
	//pView->SetScrollPos(SB_VERT, 0, TRUE);
	//pView->SetScrollPos(SB_HORZ, 0, TRUE);
	pView->RedrawWindow();

	return pages;
}


int CAddressDoc::rpSetFirstLabel(CPoint point)
// set first label for printing given a point in the view window
// +x moves from left to right, +y goes in a downward direction
// labels are printed with column major ordering
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return FALSE;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CWnd *pCBox;

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (!((CButton *)pCBox)->GetCheck())
	{
		return FALSE;
	}

	return TRUE;
}


void CAddressDoc::OnFileExportPalmCsv() 
{
	FILE *fp;
	CAddressTable *ptrTable;
	char name[256];
	char firstName[256];
	char lastName[256];
	char email[256];
	char phone[4][256];
	char street[256];
	char s[256];
	char *t;
	char *u;
	int i;
	int j;

	// TODO: Add your command handler code here
	CFileDialog fileExportCsv
		(FALSE,    // save
		NULL,    // no default extension
		NULL,    // file name
		OFN_HIDEREADONLY,
		"All files (*.*)|*.*||");

	fileExportCsv.m_ofn.lpstrTitle = "Palm: Save in comma separted value (csv) format";

	if (fileExportCsv.DoModal() == IDOK)
	{
		fp = fopen((LPCSTR)fileExportCsv.GetPathName(), "w");
		if (fp == NULL)
		{
			return;
		}

		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			// if last name first put at end of name string
			sprintf(name, "%s", ptrTable->m_name);
			u = name;
			for (t = name; *t != '\0' && *t != ','; t++);
			if (*t == ',')
			{
				*t = '\0';
				t++;
				for (; *t == ' ' && *t != '\0'; t++);
				sprintf(firstName, "\"%s\"", t);
				sprintf(lastName, "\"%s\"", u);
			}
			else 
			{
				for (; *t != ' ' && t != name; t--);
				if (*t == ' ')
				{
					*t = '\0';
					t++;
					sprintf(firstName, "\"%s\"", u);
					sprintf(lastName, "\"%s\"", t);
				}
				else  // one name entry
				{
					sprintf(firstName, "\"%s\"", u);
					sprintf(lastName, "\"\"");
				}
			}

			if (ptrTable->m_phone[0][0] != '\0')
			{
				sprintf(phone[0], "\"%s\"", ptrTable->m_phone[0]); 
			}
			else
			{
				sprintf(phone[0], "\"\"");
			}

			if (ptrTable->m_phone[1][0] != '\0')
			{
				sprintf(phone[1], "\"%s\"", ptrTable->m_phone[1]); 
			}
			else
			{
				sprintf(phone[1], "\"\"");
			}

			if (ptrTable->m_phone[2][0] != '\0')
			{
				sprintf(phone[2], "\"%s\"", ptrTable->m_phone[2]); 
			}
			else
			{
				sprintf(phone[2], "\"\"");
			}

			if (ptrTable->m_phone[3][0] != '\0')
			{
				sprintf(phone[3], "\"%s\"", ptrTable->m_phone[3]); 
			}
			else
			{
				sprintf(phone[3], "\"\"");
			}

			if (ptrTable->m_email[0][0] != '\0')
			{
				sprintf(email, "\"%s\"", ptrTable->m_email[0]); 
			}
			else
			{
				sprintf(email, "\"\"");
			}

			if (ptrTable->m_address[0][0] != '\0')
			{
				sprintf(street, "\"%s\"", ptrTable->m_address[0]);
				for (t = street; *t != '\0'; t++)
				{
					if (*t == ';')
					{
						*t = 0x0D;
						if (*(t + 1) == ' ')
						{
							*(t + 1) = 0x0A;
						}
					}
				}
			}
			else
			{
				sprintf(street, "\"\"");
			}

			// any single quote in strings need to be double quotes for palm
			j = 0;
			sprintf(s, "%s", lastName);
			for (t = s; *t != '\0'; t++)
			{
				lastName[j] = *t;
				j++;
				if (*t == '"' && t != s && *(t + 1) != '\0')
				{
					lastName[j] = '"';
					j++;
				}
			}
			lastName[j] = '\0';
			j = 0;
			sprintf(s, "%s", firstName);
			for (t = s; *t != '\0'; t++)
			{
				firstName[j] = *t;
				j++;
				if (*t == '"' && t != s && *(t + 1) != '\0')
				{
					firstName[j] = '"';
					j++;
				}
			}
			firstName[j] = '\0';
			j = 0;
			sprintf(s, "%s", email);
			for (t = s; *t != '\0'; t++)
			{
				email[j] = *t;
				j++;
				if (*t == '"' && t != s && *(t + 1) != '\0')
				{
					email[j] = '"';
					j++;
				}
			}
			email[j] = '\0';
			j = 0;
			sprintf(s, "%s", street);
			for (t = s; *t != '\0'; t++)
			{
				street[j] = *t;
				j++;
				if (*t == '"' && t != s && *(t + 1) != '\0')
				{
					street[j] = '"';
					j++;
				}
			}
			street[j] = '\0';

			for (i = 0; i < 4; i++)
			{
				j = 0;
				sprintf(s, "%s", phone[i]);
				for (t = s; *t != '\0'; t++)
				{
					phone[i][j] = *t;
					j++;
					if (*t == '"' && t != s && *(t + 1) != '\0')
					{
						phone[i][j] = '"';
						j++;
					}
				}
				phone[i][j] = '\0';
			}


			// output to file
			fprintf(fp, "%s,%s,\"\",\"\",%s,%s,%s,%s,%s,%s,\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"\n",
				lastName,
				firstName,
				phone[0],
				phone[1],
				phone[2],
				phone[3],
				email,
				street);
		}




		fclose(fp);
	}
	return;		
}


void CAddressDoc::OnFileExportCsv()
// export comma separated value file for Outlook Express 
{
	FILE *fp;
	CAddressTable *ptrTable;
	char name[256];
	char firstName[256];
	char lastName[256];
	char street[2][256];
	char note[256];
	char s[256];
	char *t;
	char *u;
	char *v;
	int i;
	int j;

	// TODO: Add your command handler code here
	CFileDialog fileExportCsv
		(FALSE,    // save
		NULL,    // no default extension
		NULL,    // file name
		OFN_HIDEREADONLY,
		"All files (*.*)|*.*||");

	fileExportCsv.m_ofn.lpstrTitle = "Outlook Express: Save in comma separted value (csv) format";

	if (fileExportCsv.DoModal() == IDOK)
	{
		fp = fopen((LPCSTR)fileExportCsv.GetPathName(), "w");
		if (fp == NULL)
		{
			return;
		}

		//fprintf(fp, "First Name,Last Name,Name,E-mail Address\n"); // Name field is empty

		fprintf(fp, "First Name,Last Name,Middle Name,Name,Nickname,E-mail Address,Home Street,Home City,Home Postal Code,Home State,Home Country/Region,Home Phone,Home Fax,Mobile Phone,Personal Web Page,Business Street,Business City,Business Postal Code,Business State,Business Country/Region,Business Phone,Business Fax,Pager,Company,Job Title,Department,Office Location,Notes\n");
		for (ptrTable = m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			// only write those entries with an email
			if (ptrTable->m_email[0][0] != '\0')
			{
				// if last name first put at end of name string
				if (m_lastNameFirst == TRUE)
				{
					sprintf(name, "%s", ptrTable->m_nameLastFirst);

					u = name;
					for (t = name; *t != '\0' && *t != ','; t++);
					if (*t == ',')
					{
						*t = '\0';
						t++;
						for (; *t == ' ' && *t != '\0'; t++);
						sprintf(firstName, "%s", t);
						sprintf(lastName, "%s", u);
					}
					else  // one name entry
					{
						sprintf(firstName, "%s", u);
						lastName[0] = '\0';
					}

				}

				else
				{
					sprintf(name, "%s", ptrTable->m_name);

					u = name;
					for (t = name; *t != '\0'; t++);

					for (; *t != ' ' && t != name; t--);
					if (*t == ' ')
					{
						for (v = t; *v == ' ' || *v == ','; v--); // remove a trailing comma
						*(v + 1) = '\0';
						t++;
						sprintf(firstName, "%s", u);
						sprintf(lastName, "%s", t);
					}
					else  // one name entry
					{
						sprintf(firstName, "%s", u);
						lastName[0] = '\0';
					}
				}

	/**
				u = name;
				for (t = name; *t != '\0' && *t != ','; t++);
				if (*t == ',')
				{
					*t = '\0';
					t++;
					for (; *t == ' ' && *t != '\0'; t++);
					sprintf(firstName, "%s", t);
					sprintf(lastName, "%s", u);
				}
				else 
				{
					for (; *t != ' ' && t != name; t--);
					if (*t == ' ')
					{
						*t = '\0';
						t++;
						sprintf(firstName, "%s", u);
						sprintf(lastName, "%s", t);
					}
					else  // one name entry
					{
						sprintf(firstName, "%s", u);
						lastName[0] = '\0';
					}
				}
	**/

				street[0][0] = '\0';
				if (ptrTable->m_address[0][0] != '\0')
				{
					j = 0;
					for (t = ptrTable->m_address[0]; *t != '\0'; t++)
					{
						if (*t == ';')
						{
							street[0][j] = '\n';
							j++;
							if (*(t + 1) == ' ')
							{
								t++;
							}
						}
						else
						{
							street[0][j] = *t;
							j++;
						}
					}
					street[0][j] = '\0';
				}
				street[1][0] = '\0';
				if (ptrTable->m_address[1][0] != '\0')
				{
					j = 0;
					for (t = ptrTable->m_address[1]; *t != '\0'; t++)
					{
						if (*t == ';')
						{
							street[1][j] = '\n';
							j++;
							if (*(t + 1) == ' ')
							{
								t++;
							}
						}
						else
						{
							street[1][j] = *t;
							j++;
						}
					}
					street[1][j] = '\0';
				}
				note[0] = '\0';
				for (i = 0; i < 9 && ptrTable->m_note[i][0] != '\0'; i++)
				{
					if (i > 0 && strlen(note) + strlen(ptrTable->m_note[i]) < 256)
					{
						sprintf(note, "%s\n%s", note, ptrTable->m_note[i]);
					}
					else if (i == 0)
					{
						sprintf(note, "%s", ptrTable->m_note[0]);
					}

				}

				// need double quote at beginning and end if string contains a comma or newline
				for (t = firstName; *t != ',' && *t != '\n' && *t != '\0'; t++);
				if (*t == ',' || *t == '\n')
				{
					sprintf(s, "%s", firstName);
					sprintf(firstName ,"\"%s\"", s);
				}
				for (t = lastName; *t != ',' && *t != '\n' && *t != '\0'; t++);
				if (*t == ',' || *t == '\n')
				{
					sprintf(s, "%s", lastName);
					sprintf(lastName ,"\"%s\"", s);
				}
				for (t = street[0]; *t != ',' && *t != '\n' && *t != '\0'; t++);
				if (*t == ',' || *t == '\n')
				{
					sprintf(s, "%s", street[0]);
					sprintf(street[0] ,"\"%s\"", s);
				}
				for (t = street[1]; *t != ',' && *t != '\n' && *t != '\0'; t++);
				if (*t == ',' || *t == '\n')
				{
					sprintf(s, "%s", street[1]);
					sprintf(street[1] ,"\"%s\"", s);
				}
				for (t = note; *t != ',' && *t != '\n' && *t != '\0'; t++);
				if (*t == ',' || *t == '\n')
				{
					sprintf(s, "%s", note);
					sprintf(note ,"\"%s\"", s);
				}

				//if (ptrTable->m_email[0][0] != '\0')
				//{
				//	fprintf(fp, "%s,%s,,%s\n", firstName, lastName, ptrTable->m_email[0]);
				//}

				fprintf(fp, "%s,%s,,,%s,%s,%s,,,,,%s,,%s,,%s,,,,,%s,%s,,,,,,%s\n",
					firstName,
					lastName,
					ptrTable->m_alias[0],
					ptrTable->m_email[0],
					street[0],
					ptrTable->m_phone[0], // Home Phone position
					ptrTable->m_phone[1], // Mobile Phone position
					street[1],
					ptrTable->m_phone[2], // Business Phone position
					ptrTable->m_phone[3], // Business Fax postion
				note);
			}
		}

		fclose(fp);
	}
	return;	
}


void CAddressDoc::OnFileMerge() 
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	FILE *fp;
	char s[256];

	// TODO: Add your command handler code here
	CFileDialog fileMerge
		(TRUE,    // open
		NULL,    // no default extension
		NULL,    // file name
		OFN_HIDEREADONLY,
		"All files (*.*)|*.*||");

	fileMerge.m_ofn.lpstrTitle = "Select a file to merge with this document";

	if (fileMerge.DoModal() == IDOK)
	{
		fp = fopen((LPCSTR)fileMerge.GetPathName(), "r");
		if (fp == NULL)
		{
			return;
		}
		else
		{
			if (rpReadDoc(fp) == FALSE)
			{
				sprintf(s, "File merge failed due to incorrect file format.");
				AfxMessageBox(s);
				fclose(fp);
				return;
			}
			else
			{
				pView->SetScrollPos(SB_VERT, 0, TRUE);
				pView->SetScrollPos(SB_HORZ, 0, TRUE);
				rpComputeDrawPositions();
				rpUpdateGroupMenu();
				SetModifiedFlag(TRUE);
			}
			fclose(fp);
		}
	}

	return;
}


void CAddressDoc::OnFileWriteSelectedToFile() 
{
	POSITION pos = GetFirstViewPosition();
	if (pos == NULL)
	{
		return;
	}
	CAddressView *pView = (CAddressView *)GetNextView(pos);
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CMainFrame *pMainFrame = (CMainFrame *)pView->GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressTable *ptrTable;
	CAddressTable *curTable;
	CAddressTable *ptrTableListTemp;
	CAddressGroup *ptrGroup;
	CAddressGroup *curGroup;
	CAddressGroup *ptrGroupListDup;

	// TODO: Add your command handler code here
	CFileDialog fileWriteSelected
		(FALSE,    // save
		NULL,    // no default extension
		NULL,    // file name
		OFN_HIDEREADONLY,
		"All files (*.*)|*.*||");

	fileWriteSelected.m_ofn.lpstrTitle = "Write selected names to a new file";

	if (fileWriteSelected.DoModal() != IDOK)
	{
		return;
	}

	// create a table list with selected names only
	ptrTableListTemp = m_ptrTableList;
	m_ptrTableList = NULL;
	for (curTable = ptrTableListTemp; curTable != NULL;)
	{
		ptrTable = curTable;
		curTable = ptrTable->m_next;

		if (ptrTable->m_selected)
		{

			if (ptrTable == ptrTableListTemp)
			// first in list
			{
				ptrTableListTemp = ptrTable->m_next;
			}
			if (ptrTable->m_next != NULL)
			{
				ptrTable->m_next->m_prev = ptrTable->m_prev;
			}
			if (ptrTable->m_prev != NULL)
			{
				ptrTable->m_prev->m_next = ptrTable->m_next;
			}
			ptrTable->m_prev = NULL;
			ptrTable->m_next = NULL;

			ptrTable->rpInsertTable(&m_ptrTableList, m_lastNameFirst);
		}
	}

	// duplicate the group list with selected names only
	ptrGroupListDup = m_ptrGroupList;
	m_ptrGroupList = NULL;
	for (curGroup = ptrGroupListDup; curGroup != NULL; curGroup = curGroup->m_next)
	{
		if (curGroup->m_ptrTable->m_selected)
		{
			ptrGroup = new CAddressGroup;
			ptrGroup->m_ptrTable = curGroup->m_ptrTable;
			sprintf(ptrGroup->m_groupName, "%s", curGroup->m_groupName);
			ptrGroup->rpInsertGroup(&m_ptrGroupList, m_lastNameFirst);
		}
	}

	rpWriteDoc(fileWriteSelected.GetPathName());

	// restore m_ptrTableList to its original state
	for (curTable = m_ptrTableList; curTable != NULL;)
	{
		ptrTable = curTable;
		curTable = ptrTable->m_next;
		ptrTable->m_prev = NULL;
		ptrTable->m_next = NULL;
		ptrTable->rpInsertTable(&ptrTableListTemp, m_lastNameFirst);
	}
	m_ptrTableList = ptrTableListTemp;

	// restore m_ptrGroupList to its original state
	for (curGroup = m_ptrGroupList; curGroup != NULL;)
	{
		ptrGroup = curGroup;
		curGroup = ptrGroup->m_next;
		delete ptrGroup;
	}
	m_ptrGroupList = ptrGroupListDup;

	return;	
}

void CAddressDoc::OnUpdateFileWriteSelectedToFile(CCmdUI* pCmdUI) 
{
	// TODO: Add your command update UI handler code here
	CAddressTable *curTable;

	pCmdUI->Enable(FALSE);

	for (curTable = m_ptrTableList; curTable != NULL; curTable = curTable->m_next)
	{
		if (curTable->m_selected)
		{
			pCmdUI->Enable(TRUE);
		}
	}

	return;
}
