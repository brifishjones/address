// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "address.h"
#include "addressAddEntryDlg.h"
#include "addressReturnAddressDlg.h"
#include "addressGroupDlg.h"
#include "addressDoc.h"
#include "addressView.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWnd)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
	//{{AFX_MSG_MAP(CMainFrame)
	ON_WM_CREATE()
	ON_COMMAND(ID_VIEW_RETURN_ADDRESS, OnViewReturnAddress)
	ON_COMMAND(ID_FONT_CHOOSE, OnFontChoose)
	ON_COMMAND(ID_FONT_RESTORE_DEFAULT, OnFontRestoreDefault)
	ON_COMMAND(ID_GROUP_UPDATE, OnGroupUpdate)
	ON_UPDATE_COMMAND_UI(ID_GROUP_UPDATE, OnUpdateGroupUpdate)
	ON_UPDATE_COMMAND_UI(ID_GROUP_ONE, OnUpdateGroupOne)
	ON_UPDATE_COMMAND_UI(ID_GROUP_FIVE, OnUpdateGroupFive)
	ON_UPDATE_COMMAND_UI(ID_GROUP_FOUR, OnUpdateGroupFour)
	ON_UPDATE_COMMAND_UI(ID_GROUP_SEVEN, OnUpdateGroupSeven)
	ON_UPDATE_COMMAND_UI(ID_GROUP_SIX, OnUpdateGroupSix)
	ON_UPDATE_COMMAND_UI(ID_GROUP_THREE, OnUpdateGroupThree)
	ON_UPDATE_COMMAND_UI(ID_GROUP_TWO, OnUpdateGroupTwo)
	ON_COMMAND(ID_GROUP_FIVE, OnGroupFive)
	ON_COMMAND(ID_GROUP_FOUR, OnGroupFour)
	ON_COMMAND(ID_GROUP_ONE, OnGroupOne)
	ON_COMMAND(ID_GROUP_SEVEN, OnGroupSeven)
	ON_COMMAND(ID_GROUP_SIX, OnGroupSix)
	ON_COMMAND(ID_GROUP_THREE, OnGroupThree)
	ON_COMMAND(ID_GROUP_TWO, OnGroupTwo)
	ON_COMMAND(ID_VIEW_LAST_NAME_FIRST, OnViewLastNameFirst)
	ON_UPDATE_COMMAND_UI(ID_VIEW_LAST_NAME_FIRST, OnUpdateViewLastNameFirst)
	ON_COMMAND(ID_GROUP_EIGHT, OnGroupEight)
	ON_UPDATE_COMMAND_UI(ID_GROUP_EIGHT, OnUpdateGroupEight)
	ON_COMMAND(ID_GROUP_NINE, OnGroupNine)
	ON_UPDATE_COMMAND_UI(ID_GROUP_NINE, OnUpdateGroupNine)
	ON_COMMAND(ID_GROUP_TEN, OnGroupTen)
	ON_UPDATE_COMMAND_UI(ID_GROUP_TEN, OnUpdateGroupTen)
	ON_COMMAND(ID_GROUP_ELEVEN, OnGroupEleven)
	ON_UPDATE_COMMAND_UI(ID_GROUP_ELEVEN, OnUpdateGroupEleven)
	ON_COMMAND(ID_GROUP_TWELVE, OnGroupTwelve)
	ON_UPDATE_COMMAND_UI(ID_GROUP_TWELVE, OnUpdateGroupTwelve)
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_COPY, OnUpdateEditCopy)
	//}}AFX_MSG_MAP
	ON_COMMAND(IDC_LIST, OnList)
	ON_COMMAND(IDC_BOOK, OnBook)
	ON_COMMAND(IDC_ENVELOPE, OnEnvelope)
	ON_COMMAND(IDC_LABEL, OnLabel)
	ON_COMMAND(IDC_RETURN, OnReturn)
	ON_COMMAND(IDC_LINE_SPACING_DEFAULTS, OnLineSpacingDefaults)
	ON_COMMAND(IDC_COLUMN_WIDTH_DEFAULTS, OnColumnWidthDefaults)
	ON_COMMAND(IDC_SELECT_ALL, OnSelectAll)
	ON_COMMAND(IDC_SELECT_NONE, OnSelectNone)
	ON_COMMAND(IDC_ADD, OnAddEntry)
	ON_COMMAND(IDC_DELETE, OnDelete)
	ON_LBN_SELCHANGE(IDC_NAME, OnName)
	ON_LBN_DBLCLK(IDC_NAME, OnNameEditEntry)
	ON_UPDATE_COMMAND_UI(IDC_LIST, OnUpdateListUI)
	ON_UPDATE_COMMAND_UI(IDC_BOOK, OnUpdateBookUI)
	ON_UPDATE_COMMAND_UI(IDC_ENVELOPE, OnUpdateEnvelopeUI)
	ON_UPDATE_COMMAND_UI(IDC_LABEL, OnUpdateLabelUI)
	ON_EN_KILLFOCUS(IDC_LINE_SPACING, OnLineSpacing)
	ON_EN_KILLFOCUS(IDC_COLUMN_WIDTH, OnColumnWidth)
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	//ID_INDICATOR_CAPS,
	//ID_INDICATOR_NUM,
	//ID_INDICATOR_SCRL,
};

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	// TODO: add member initialization code here
	
}

CMainFrame::~CMainFrame()
// free memory for workcell and program nodes
{
	/***
	CAddressAlias *ptrWorkcell;
	CAddressAlias *curWorkcell;
	CAddressProgram *ptrProgram;
	CAddressProgram *curProgram;

	for (curWorkcell = m_ptrWorkcellList; curWorkcell != NULL;)
	{
		for (curProgram = curWorkcell->m_ptrProgram; curProgram != NULL;)
		{
			ptrProgram = curProgram;
			curProgram = curProgram->m_next;
			delete ptrProgram;
		}

		ptrWorkcell = curWorkcell;
		curWorkcell = curWorkcell->m_next;
		delete ptrWorkcell;
	}
	***/
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
// create and initialize main frame and top dialog bar
{
	CWnd *pCBox;
	CMenu *menu;

	if (CFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
		  sizeof(indicators)/sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	if (!m_wndDlgBarTop.Create(this, IDD_ADDRESS_TOP,
		CBRS_TOP|CBRS_TOOLTIPS|CBRS_FLYBY, IDD_ADDRESS_TOP))  
	{
		TRACE0("Failed to create top dialog bar\n");
		return -1;      // fail to create
	}

	if (!m_wndDlgBarLeft.Create(this, IDD_ADDRESS_LEFT,
		CBRS_LEFT|CBRS_TOOLTIPS|CBRS_FLYBY, IDD_ADDRESS_LEFT))  
	{
		TRACE0("Failed to create left dialog bar\n");
		return -1;      // fail to create
	}

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	((CButton *)pCBox)->SetCheck(TRUE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	((CButton *)pCBox)->SetCheck(FALSE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH_DEFAULTS);
	pCBox->EnableWindow(FALSE);


	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_SELECT_ALL);
	pCBox->EnableWindow(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_SELECT_NONE);
	pCBox->EnableWindow(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_ADD);
	pCBox->EnableWindow(TRUE);
	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_DELETE);
	pCBox->EnableWindow(FALSE);

	// initialize group menu
	if ((menu = this->GetMenu()) != NULL)
	{
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

		menu->ModifyMenu(ID_GROUP_TWELVE, MF_BYCOMMAND, ID_GROUP_TWELVE, "(empty)");
	}
	
	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWnd::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWnd::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWnd::Dump(dc);
}

#endif //_DEBUG


void CMainFrame::rpCopyToClipboard(char *s)
{
	// copy info to the clipboard
	HGLOBAL hMem;
	LPTSTR pszDst;
	hMem = GlobalAlloc(GHND, (DWORD)(lstrlen(s) + 1));

    if (hMem != NULL)
	{        
		pszDst = (char *)GlobalLock(hMem);
		lstrcpy(pszDst, s);
		GlobalUnlock(hMem);

		if (OpenClipboard())
		{
			TRACE1("CMainFrame::rpCopyToClipboard()   %ld chars copied\n", lstrlen(s));
			EmptyClipboard();
			SetClipboardData(CF_TEXT, hMem);
			CloseClipboard();
		}
		else
		{
			GlobalFree(hMem);
		}
	}
	
	return;
}



void CMainFrame::rpReadEntryDlg(CAddressTable *ptrTable, CAddressAddEntryDlg *addEntry)
// takes information entered in dialog box addEntry and puts it into ptrTable
// if entry is being added then ptrTable is NULL
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	int i;
	int j;
	char *t;
	char *u;
	char s[2056];
	char entry[256];
	char alias[256];
	BOOL validName = FALSE;
	
	// read name
	sprintf(s, "%s", (LPCTSTR)addEntry->m_name);
	i = 0;
	for (t = s; *t == ' '; t++);  // remove leading white space if any
	for (; *t != '\0'; t++)
	{
		if (*t == 0x0D && *(t + 1) == 0x0A)
		{
			entry[i] = ' ';
			t++;
		}
		else
		{
			entry[i] = *t;
		}
		i++;
		if (validName == FALSE && *t >= 48 && *t <= 165)
		{
			validName = TRUE;
		}
	}
	entry[i] = '\0';

	// remove any trailing whitespace or semi-colins
	for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
	{
		*t = '\0';
	}

	//pDoc->rpSortName(entry);
	//addEntry->m_name = entry;

	if (validName == FALSE)
	{
		return;
	}
	else if (ptrTable != NULL)
	{
		if (ptrTable == pDoc->m_ptrTableList)
		// first in list
		{
			pDoc->m_ptrTableList = ptrTable->m_next;
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
		ptrTable->m_numberOfLinesBook = 0;
		ptrTable->m_numberOfLinesLabel = 0;
		ptrTable->m_selected = FALSE;
		ptrTable->m_drawCoord[0] = 0;
		ptrTable->m_drawCoord[1] = 0;
		ptrTable->m_drawCoord[2] = 0;
		ptrTable->m_drawCoord[3] = 0;
		ptrTable->m_name[0] = '\0';
		for (j = 0; j < 10; j++)
		{
			ptrTable->m_email[j][0] = '\0';
			ptrTable->m_alias[j][0] = '\0';
			ptrTable->m_address[j][0] = '\0';
			ptrTable->m_phone[j][0] = '\0';
			ptrTable->m_note[j][0] = '\0';
		}
		//delete ptrTable;
	}
	else
	{
		// create a new table
		ptrTable = new CAddressTable;
		// by default make new entry visible in main view window
		ptrTable->m_selected = TRUE;
	}

	sprintf(ptrTable->m_name, "%s", entry);
	ptrTable->m_numberOfLinesBook++;
	ptrTable->m_numberOfLinesLabel++;

	// read address
	sprintf(s, "%s", (LPCTSTR)addEntry->m_address);
	if (s[0] != '\0')
	{
		i = 0;
		// remove leading white space and line feeds if any
		for (t = s; *t == ' ' || *t == 0x0D || *t == 0x0A; t++);
		for (; *t != '\0'; t++)
		{
			if (*t == 0x0D && *(t + 1) == 0x0A)
			{
				entry[i] = ';';
				entry[i + 1] = ' ';
				*(t + 1) = ' ';
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				entry[i] = *t;
			}
			i++;
		}
		entry[i] = '\0';

		// remove any trailing whitespace or semi-colins
		for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
		{
			*t = '\0';
		}

		j = 0;
		u = entry;
		for (t = entry; *t != '\0'; t++)
		{
			if (*t == ';' && *(t + 1) == ' ' && *(t + 2) == ';')
			{
				*t = '\0';
				sprintf(ptrTable->m_address[j], "%s", u);
				ptrTable->m_numberOfLinesBook++;
				if (j == 0)
				{
					ptrTable->m_numberOfLinesLabel++;
				}
				t++;
				while(*t == ' ' && *(t + 1) == ';' || *t == ';' && *(t + 1) == ' ')
				{
					t++;
				}
				j++;
				u = t + 1;
			}
			else if (*t == ';' && *(t + 1) == ' ')
			{
				ptrTable->m_numberOfLinesBook++;
				if (j == 0)
				{
					ptrTable->m_numberOfLinesLabel++;
				}
			}
		}
		sprintf(ptrTable->m_address[j], "%s", u);
		ptrTable->m_numberOfLinesBook++;
		if (j == 0)
		{
			ptrTable->m_numberOfLinesLabel++;
		}
	}

	// read phone
	sprintf(s, "%s", (LPCTSTR)addEntry->m_phone);
	if (s[0] != '\0')
	{
		i = 0;
		// remove leading white space and line feeds if any
		for (t = s; *t == ' ' || *t == 0x0D || *t == 0x0A; t++);
		for (; *t != '\0'; t++)
		{
			if (*t == 0x0D && *(t + 1) == 0x0A)
			{
				entry[i] = ';';
				entry[i + 1] = ' ';
				*(t + 1) = ' ';
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				entry[i] = *t;
			}
			i++;
		}
		entry[i] = '\0';

		// remove any trailing whitespace or semi-colins
		for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
		{
			*t = '\0';
		}

		j = 0;
		u = entry;
		for (t = entry; *t != '\0'; t++)
		{
			if (*t == ';' && *(t + 1) == ' ')
			{
				*t = '\0';
				sprintf(ptrTable->m_phone[j], "%s", u);
				ptrTable->m_numberOfLinesBook++;
				t++;
				while(*t == ' ' && *(t + 1) == ';' || *t == ';' && *(t + 1) == ' ')
				{
					t++;
				}
				j++;
				u = t + 1;
			}
		}
		sprintf(ptrTable->m_phone[j], "%s", u);
		ptrTable->m_numberOfLinesBook++;
	}

	// read notes
	sprintf(s, "%s", (LPCTSTR)addEntry->m_note);
	if (s[0] != '\0')
	{
		i = 0;
		// remove leading white space and line feeds if any
		for (t = s; *t == ' ' || *t == 0x0D || *t == 0x0A; t++);
		for (; *t != '\0'; t++)
		{
			if (*t == 0x0D && *(t + 1) == 0x0A)
			{
				entry[i] = ';';
				entry[i + 1] = ' ';
				*(t + 1) = ' ';
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				entry[i] = *t;
			}
			i++;
		}
		entry[i] = '\0';

		// remove any trailing whitespace or semi-colins
		for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
		{
			*t = '\0';
		}

		j = 0;
		u = entry;
		for (t = entry; *t != '\0'; t++)
		{
			if (*t == ';' && *(t + 1) == ' ')
			{
				*t = '\0';
				sprintf(ptrTable->m_note[j], "%s", u);
				ptrTable->m_numberOfLinesBook++;
				t++;
				while(*t == ' ' && *(t + 1) == ';' || *t == ';' && *(t + 1) == ' ')
				{
					t++;
				}
				j++;
				u = t + 1;
			}
		}
		sprintf(ptrTable->m_note[j], "%s", u);
		ptrTable->m_numberOfLinesBook++;
	}

	// read email
	sprintf(s, "%s", (LPCTSTR)addEntry->m_email);
	if (s[0] != '\0')
	{
		i = 0;
		// remove leading white space and line feeds if any
		for (t = s; *t == ' ' || *t == 0x0D || *t == 0x0A; t++);
		for (; *t != '\0'; t++)
		{
			if (*t == 0x0D && *(t + 1) == 0x0A)
			{
				entry[i] = ';';
				entry[i + 1] = ' ';
				*(t + 1) = ' ';
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				entry[i] = *t;
			}
			i++;
		}
		entry[i] = '\0';

		// remove any trailing whitespace or semi-colins
		for (t = &entry[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
		{
			*t = '\0';
		}

		j = 0;
		u = entry;
		for (t = entry; *t != '\0'; t++)
		{
			if (*t == ';' && *(t + 1) == ' ')
			{
				*t = '\0';
				i = sscanf(u, "%s %s", &ptrTable->m_email[j], alias);
				ptrTable->m_numberOfLinesBook++;
				if (i == 2)
				{
					i = 0;
					for (u = alias; *u != '\0'; u++)
					{
						if (*u != '(' && *u != ')' && *u != '[' && *u != ']')
						{
							ptrTable->m_alias[j][i] = *u;
							i++;
						}
					}
					ptrTable->m_alias[j][i] = '\0';
				}

				t++;
				while(*t == ' ' && *(t + 1) == ';' || *t == ';' && *(t + 1) == ' ')
				{
					t++;
				}
				j++;
				u = t + 1;
			}
		}
		i = sscanf(u, "%s %s", &ptrTable->m_email[j], alias);
		ptrTable->m_numberOfLinesBook++;
		if (i == 2)
		{
			i = 0;
			for (u = alias; *u != '\0'; u++)
			{
				if (*u != '(' && *u != ')' && *u != '[' && *u != ']')
				{
					ptrTable->m_alias[j][i] = *u;
					i++;
				}
			}
			ptrTable->m_alias[j][i] = '\0';
		}
	}

	ptrTable->m_selected = TRUE;
	ptrTable->rpInsertTable(&pDoc->m_ptrTableList, pDoc->m_lastNameFirst);

	return;
}

void CMainFrame::rpWriteEntryDlg(CAddressTable *ptrTable, CAddressAddEntryDlg *addEntry)
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	int i;
	int j;
	char *t;
	char s[2056];
	BOOL validName = FALSE;

	// TODO: Add your control notification handler code here
	if (ptrTable == NULL)
	{
		return;
	}

	// name
	addEntry->m_name.Format("%s", (LPCTSTR)ptrTable->m_name);

	// address
	i = 0;
	for (j = 0; j < 9 && ptrTable->m_address[j][0] != '\0'; j++)
	{
		for (t = ptrTable->m_address[j]; *t == ' '; t++);  // remove leading white space if any
		for (; *t != '\0'; t++)
		{
			if (*t == ';')
			{
				// cannot use \n for newlines in edit control
				// rather CR LF character combination (0x0D 0x0A)
				s[i] = 0x0D;
				s[i + 1] = 0x0A;
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				s[i] = *t;
			}
			i++;
		}
		if (j < 8 && ptrTable->m_address[j + 1][0] != '\0')
		{
			s[i] = 0x0D;
			s[i + 1] = 0x0A;
			s[i + 2] = 0x0D;
			s[i + 3] = 0x0A;
			i += 4;
		}

	}
	s[i] = '\0';
	addEntry->m_address.Format("%s", (LPCTSTR)s);

	// phone
	i = 0;
	for (j = 0; j < 9 && ptrTable->m_phone[j][0] != '\0'; j++)
	{
		for (t = ptrTable->m_phone[j]; *t == ' '; t++);  // remove leading white space if any
		for (; *t != '\0'; t++)
		{
			s[i] = *t;
			i++;
		}
		if (j < 8 && ptrTable->m_phone[j + 1][0] != '\0')
		{
			s[i] = 0x0D;
			s[i + 1] = 0x0A;
			i += 2;
		}
	}
	s[i] = '\0';
	addEntry->m_phone.Format("%s", (LPCTSTR)s);

	// note
	i = 0;
	for (j = 0; j < 9 && ptrTable->m_note[j][0] != '\0'; j++)
	{
		for (t = ptrTable->m_note[j]; *t == ' '; t++);  // remove leading white space if any
		for (; *t != '\0'; t++)
		{
			s[i] = *t;
			i++;
		}
		if (j < 8 && ptrTable->m_note[j + 1][0] != '\0')
		{
			s[i] = 0x0D;
			s[i + 1] = 0x0A;
			i += 2;
		}
	}
	s[i] = '\0';
	addEntry->m_note.Format("%s", (LPCTSTR)s);

	// email (alias)
	i = 0;
	for (j = 0; j < 9 && ptrTable->m_email[j][0] != '\0'; j++)
	{
		for (t = ptrTable->m_email[j]; *t == ' '; t++);  // remove leading white space if any
		for (; *t != '\0'; t++)
		{
			s[i] = *t;
			i++;
		}

		if (ptrTable->m_alias[j][0] != '\0')
		{
			s[i] = ' ';
			s[i + 1] = '(';
			i += 2;
			for (t = ptrTable->m_alias[j]; *t == ' '; t++);  // remove leading white space if any
			for (; *t != '\0'; t++)
			{
				s[i] = *t;
				i++;
			}
			s[i] = ')';
			i++;
		}

		if (j < 8 && ptrTable->m_email[j + 1][0] != '\0')
		{
			s[i] = 0x0D;
			s[i + 1] = 0x0A;
			i += 2;
		}
	}
	s[i] = '\0';
	addEntry->m_email.Format("%s", (LPCTSTR)s);

	return;
}

BOOL CMainFrame::rpSelectGroupMembers(int groupID)
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return FALSE;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return FALSE;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	int k;
	int numberOfGroups;

	numberOfGroups = 0;
	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == pDoc->m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			numberOfGroups++;
		}
	}

	k = 0;
	i = 0;
	ptrTable = pDoc->m_ptrTableList;
	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == pDoc->m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			m_allGroupMembersSelected[k] = TRUE;
			k++;
		}
		if (k == groupID || k == numberOfGroups && groupID == MAXIMUM_NUMBER_GROUPS)
		{
			for (; ptrTable != NULL && ptrTable != ptrGroup->m_ptrTable; ptrTable = ptrTable->m_next)
			{
				i++;
			}
			if (ptrTable != NULL)
			{
				ptrTable->m_selected = TRUE;
				i++;
				ptrTable = ptrTable->m_next;
			}
		}
		if (m_allGroupMembersSelected[k - 1] == TRUE && ptrGroup->m_ptrTable->m_selected == FALSE)
		{
			m_allGroupMembersSelected[k - 1] = FALSE;
		}
	}

	pDoc->rpComputeDrawPositions();

	return TRUE;
}


BOOL CMainFrame::rpUpdateGroupMenu()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		// fails when address is executed by opening a document from the desktop
		return FALSE;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return FALSE;
	}
	CMenu *menu;
	char groupName[30][64];
	int k = 0;
	int i;
	static int prevGroupCount = 30;
	CAddressGroup *ptrGroup;

	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == pDoc->m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			m_allGroupMembersSelected[k] = TRUE;
			sprintf(groupName[k], "%s", ptrGroup->m_groupName);
			k++;
		}
		if (m_allGroupMembersSelected[k - 1] == TRUE && ptrGroup->m_ptrTable->m_selected == FALSE)
		{
			m_allGroupMembersSelected[k - 1] = FALSE;
		}
	}

	if ((menu = this->GetMenu()) == NULL)
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



/////////////////////////////////////////////////////////////////////////////
// CMainFrame message handlers

void CMainFrame::OnList()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	char s[32];
	int i;
	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	((CButton *)pCBox)->SetCheck(TRUE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	((CButton *)pCBox)->SetCheck(FALSE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING_DEFAULTS);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH_DEFAULTS);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
	pCBox->ShowWindow(SW_HIDE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	sprintf(s, "%d", pDoc->m_lineSpacingList);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_SELECT_ALL);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
		i = ((CListBox *)pCBox)->GetCount() - 1;
		if (i > 0)
		{
			((CListBox *)pCBox)->SelItemRange(TRUE, 0, i);
		}
		else
		{
			((CListBox *)pCBox)->SetSel(0, TRUE);
		}
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	pDoc->rpComputeDrawPositions();

	return;
}


void CMainFrame::OnBook()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	char s[32];
	int i;
	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	((CButton *)pCBox)->SetCheck(TRUE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	((CButton *)pCBox)->SetCheck(FALSE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING_DEFAULTS);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH_DEFAULTS);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
	pCBox->ShowWindow(SW_HIDE);


	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	sprintf(s, "%d", pDoc->m_lineSpacing);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_SELECT_ALL);
	if (((CButton *)pCBox)->GetCheck())
	{
		pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
		i = ((CListBox *)pCBox)->GetCount() - 1;
		if (i > 0)
		{
			((CListBox *)pCBox)->SelItemRange(TRUE, 0, i);
		}
		else
		{
			((CListBox *)pCBox)->SetSel(0, TRUE);
		}
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);
	
	pDoc->rpComputeDrawPositions();

	return;
}


void CMainFrame::OnEnvelope()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	char s[32];
	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	((CButton *)pCBox)->SetCheck(TRUE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	((CButton *)pCBox)->SetCheck(FALSE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING_DEFAULTS);
	pCBox->ShowWindow(SW_SHOW);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH_DEFAULTS);
	pCBox->ShowWindow(SW_HIDE);
	//pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
	//pCBox->ShowWindow(SW_HIDE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	sprintf(s, "%d", pDoc->m_lineSpacingEnvelope);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	if (pDoc->m_returnAddress[0] == '\0')
	{
		pDoc->m_return = FALSE;
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		pCBox->ShowWindow(SW_HIDE);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
		pCBox->ShowWindow(SW_HIDE);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
		pCBox->ShowWindow(SW_HIDE);
	}
	else
	{
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		pCBox->ShowWindow(SW_SHOW);
		if (((CButton *)pCBox)->GetCheck())
		{
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_SHOW);
			sprintf(s, "%d", pDoc->m_returnSizeEnvelope);
			pCBox->SetWindowText(s);
			pCBox->UpdateWindow();
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_SHOW);
		}
		else
		{
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	pDoc->rpComputeDrawPositions();

	return;
}



void CMainFrame::OnLabel()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	((CButton *)pCBox)->SetCheck(FALSE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	((CButton *)pCBox)->SetCheck(TRUE);

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING_DEFAULTS);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	pCBox->ShowWindow(SW_HIDE);
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH_DEFAULTS);
	pCBox->ShowWindow(SW_HIDE);
	if (pDoc->m_returnAddress[0] == '\0')
	{
		pDoc->m_return = FALSE;
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		pCBox->ShowWindow(SW_HIDE);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
		pCBox->ShowWindow(SW_HIDE);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
		pCBox->ShowWindow(SW_HIDE);
	}
	else
	{
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
		pCBox->ShowWindow(SW_SHOW);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
		pCBox->ShowWindow(SW_HIDE);
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
		pCBox->ShowWindow(SW_HIDE);
	}

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	pDoc->rpComputeDrawPositions();
	return;
}

void CMainFrame::OnName()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressTable *ptrTable;
	CAddressGroup *ptrGroup;
	int i;
	int k;
	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	k = ((CListBox *)pCBox)->GetCaretIndex();

	i = 0;
	for (ptrTable = pDoc->m_ptrTableList;
		ptrTable != NULL && i != k;
		ptrTable = ptrTable->m_next)
	{
		i++;
	}

	if (ptrTable != NULL)
	{
		if (((CListBox *)pCBox)->GetSel(k))
		{
			ptrTable->m_selected = TRUE;
		}
		else
		{
			ptrTable->m_selected = FALSE;
		}
	}

	// update group menu member selection status
	k = 0;
	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == pDoc->m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			m_allGroupMembersSelected[k] = TRUE;
			k++;
		}
		if (m_allGroupMembersSelected[k - 1] == TRUE && ptrGroup->m_ptrTable->m_selected == FALSE)
		{
			m_allGroupMembersSelected[k - 1] = FALSE;
		}
	}

	pDoc->rpComputeDrawPositions();

	return;
}


void CMainFrame::OnNameEditEntry()
// called when an entry in IDC_NAME list box is double clicked with left mouse button
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressTable *ptrTable;
	CAddressTable *curTable;
	CAddressTable *prev;
	CAddressTable *next;
	CAddressAddEntryDlg addEntry;
	int i;
	int k;
	char s[256];
	CWnd *pCBox;

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	k = ((CListBox *)pCBox)->GetCaretIndex();
	((CListBox *)pCBox)->SetSel(k, TRUE);

	i = 0;
	for (ptrTable = pDoc->m_ptrTableList;
		ptrTable != NULL && i != k;
		ptrTable = ptrTable->m_next)
	{
		i++;
	}

	rpWriteEntryDlg(ptrTable, &addEntry);
	sprintf(s, "%s", ptrTable->m_name);
	prev = ptrTable->m_prev;
	next = ptrTable->m_next;
	pCBox = (CComboBox*)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);

	addEntry.m_title = "Edit entry";

	if (addEntry.DoModal() == IDOK)
	{
		rpReadEntryDlg(ptrTable, &addEntry);
		//if (strcmp(s, ptrTable->m_name) != 0)
		//{
		//	((CListBox *)pCBox)->DeleteString(k);

		//	if (ptrTable->m_prev == prev && ptrTable->m_next == next)
		//	{
		//		((CListBox *)pCBox)->InsertString(k, ptrTable->m_name);
		//		ptrTable->m_selected = TRUE;
		//		//((CListBox *)pCBox)->SetSel(k, TRUE);
		//	}
		//	else
		//	{
		//		i = 0;
		//		for (curTable = pDoc->m_ptrTableList;
		//			curTable != NULL
		//			&& strcmp(ptrTable->m_name, curTable->m_name) > 0;
		//			curTable = curTable->m_next)
		//		{
		//			i++;
		//		}
		//		((CListBox *)pCBox)->InsertString(i, ptrTable->m_name);
		//		ptrTable->m_selected = TRUE;
		//		//((CListBox *)pCBox)->SetSel(i, TRUE);
		//	}
		//}
		//else
		//{
		//	ptrTable->m_selected = TRUE;
		//	//((CListBox *)pCBox)->SetSel(k, TRUE);
		//}

		// if the name has changed the group list may need to be reordered
		pDoc->rpSortName();
		pDoc->SetModifiedFlag(TRUE);
	}

	else
	{
		//((CListBox *)pCBox)->SetSel(k, TRUE);
		ptrTable->m_selected = TRUE;
		pDoc->rpComputeDrawPositions();
		pView->RedrawWindow();
	}

	return;
}

void CMainFrame::OnAddEntry()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressAddEntryDlg addEntry;
	CAddressTable *curTable;
	CAddressGroup *ptrGroup;
	CWnd *pCBox;
	int i;
	int k;
	char s[256];

	// TODO: Add your control notification handler code here
	addEntry.m_title = "Add entry";

	if (addEntry.DoModal() == IDOK)
	{
		rpReadEntryDlg(NULL, &addEntry);
		pDoc->rpUpdateAddressList();

		//i = 0;
		//for (curTable = pDoc->m_ptrTableList;
		//	curTable != NULL
		//	&& strcmp((LPCTSTR)addEntry.m_name, curTable->m_name) > 0;
		//	curTable = curTable->m_next)
		//{
		//	i++;
		//}

		//pCBox = (CComboBox*)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
		//((CListBox *)pCBox)->GetText(i, s);
		//if (strcmp(curTable->m_name, s) != 0)
		//{
		//	// addEntry is not a duplicate entry
		//	((CListBox *)pCBox)->InsertString(i, addEntry.m_name);
		//}
		//((CListBox *)pCBox)->SetSel(i, TRUE);
		//curTable->m_selected = TRUE;

		pDoc->rpComputeDrawPositions();
		pDoc->SetModifiedFlag(TRUE);
	}

	// update group menu member selection status
	k = 0;
	for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		if (ptrGroup == pDoc->m_ptrGroupList
			|| (ptrGroup->m_prev != NULL
			&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		{
			m_allGroupMembersSelected[k] = TRUE;
			k++;
		}
		if (m_allGroupMembersSelected[k - 1] == TRUE && ptrGroup->m_ptrTable->m_selected == FALSE)
		{
			m_allGroupMembersSelected[k - 1] = FALSE;
		}
	}

	return;
}

void CMainFrame::OnDelete()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CWnd *pCBox;
	CAddressGroup *ptrGroup;
	CAddressGroup *prevGroup;
	CAddressTable *ptrTable;
	CAddressTable *prevTable;
	char s[128];
	int k;
	int i;

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_DELETE);
	((CButton *)pCBox)->SetCheck(FALSE);

	pCBox = (CComboBox*)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	k = 0;
	i = 0;	
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (((CListBox *)pCBox)->GetSel(k))
		{
			i++;
		}
		k++;
	}

	sprintf(s, "Permanently remove the %d selected items?", i);
	if (AfxMessageBox(s, MB_OKCANCEL, NULL) == IDOK)
	{
		k = 0;
		for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL;)
		{
			if (((CListBox *)pCBox)->GetSel(k))
			{
				if (ptrTable == pDoc->m_ptrTableList)
				// first in list
				{
					pDoc->m_ptrTableList = ptrTable->m_next;
				}
				if (ptrTable->m_next != NULL)
				{
					ptrTable->m_next->m_prev = ptrTable->m_prev;
				}
				if (ptrTable->m_prev != NULL)
				{
					ptrTable->m_prev->m_next = ptrTable->m_next;
				}

				// remove entries from group list for this table entry
				for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL;)
				{
					if (ptrGroup->m_ptrTable == ptrTable)
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

				prevTable = ptrTable;
				ptrTable = ptrTable->m_next;
				delete prevTable;
			}
			else
			{
				ptrTable = ptrTable->m_next;
			}
			k++;
		}

		pDoc->rpUpdateAddressList();
		pDoc->rpComputeDrawPositions();
		pDoc->SetModifiedFlag(TRUE);

		pDoc->rpUpdateGroupMenu();
	}

	return;
}

void CMainFrame::OnSelectAll()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CWnd *pCBox;
	CAddressTable *curTable;
	int i;

	for (curTable = pDoc->m_ptrTableList; curTable != NULL; curTable = curTable->m_next)
	{
		curTable->m_selected = TRUE;
	}

	//pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_SELECT_ALL);
	//if (((CButton *)pCBox)->GetCheck())
	//{
	//	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	//	if (((CButton *)pCBox)->GetCheck())
	//	{
			pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
			i = ((CListBox *)pCBox)->GetCount() - 1;
			if (i > 0)
			{
				((CListBox *)pCBox)->SelItemRange(TRUE, 0, i);
			}
			else
			{
				((CListBox *)pCBox)->SetSel(0, TRUE);
			}
			//((CListBox *)pCBox)->SelItemRange(TRUE, 0, ((CListBox *)pCBox)->GetCount() - 1);
	//	}
	//	else
	//	{
	//		pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	//		((CListBox *)pCBox)->SelItemRange(TRUE, 0, ((CListBox *)pCBox)->GetCount() - 1);
	//	}
	//}
	//else
	//{
	//	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	//	((CListBox *)pCBox)->SelItemRange(FALSE, 0, ((CListBox *)pCBox)->GetCount() - 1);
	//}

	for (i = 0; i < 30; i++)
	{
		m_allGroupMembersSelected[i] = TRUE;
	}

	pDoc->rpComputeDrawPositions();

	return;
}


void CMainFrame::OnSelectNone()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	int i;
	CWnd *pCBox;
	CAddressTable *curTable;

	for (curTable = pDoc->m_ptrTableList; curTable != NULL; curTable = curTable->m_next)
	{
		curTable->m_selected = FALSE;
	}

	pCBox = (CComboBox *)m_wndDlgBarLeft.GetDlgItem(IDC_NAME);
	((CListBox *)pCBox)->SelItemRange(FALSE, 0, ((CListBox *)pCBox)->GetCount() - 1);

	for (i = 0; i < 30; i++)
	{
		m_allGroupMembersSelected[i] = FALSE;
	}

	pDoc->m_firstLabel = 0;
	pDoc->m_lastLabel = 0;

	pView->SetScrollPos(SB_VERT, 0, TRUE);
	pView->SetScrollPos(SB_HORZ, 0, TRUE);

	pDoc->rpComputeDrawPositions();


	return;
}


void CMainFrame::OnLineSpacing()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[32];
	int i;
	int lineSpacing;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->GetWindowText(s, 32);
	i = sscanf(s, "%d", &lineSpacing);
	
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		if (i == 1)
		{
			lineSpacing = MAX(lineSpacing, 8);
			lineSpacing = MIN(lineSpacing, 72);
			if (pDoc->m_lineSpacingEnvelope != lineSpacing)
			{
				pDoc->m_lineSpacingEnvelope = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
			}
		}
		sprintf(s, "%d", pDoc->m_lineSpacingEnvelope);
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		if (i == 1)
		{
			lineSpacing = MAX(lineSpacing, 8);
			lineSpacing = MIN(lineSpacing, 72);
			if (pDoc->m_lineSpacingList != lineSpacing)
			{
				pDoc->m_lineSpacingList = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
			}
		}
		sprintf(s, "%d", pDoc->m_lineSpacingList);
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		{
			lineSpacing = MAX(lineSpacing, 8);
			lineSpacing = MIN(lineSpacing, 72);
			if (pDoc->m_lineSpacing != lineSpacing)
			{
				pDoc->m_lineSpacing = lineSpacing;
				pDoc->rpComputeDrawPositions();
				if (pDoc->m_ptrTableList != NULL)
				{
					pDoc->SetModifiedFlag(TRUE);
				}
			}
		}
		sprintf(s, "%d", pDoc->m_lineSpacing);
	}

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	return;
}

void CMainFrame::OnColumnWidth()

{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[32];
	int i;
	int columnWidth;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	pCBox->GetWindowText(s, 32);
	i = sscanf(s, "%d", &columnWidth);
	
	if (i == 1)
	{
		columnWidth = MAX(columnWidth, 100);
		columnWidth = MIN(columnWidth,
			pDoc->m_pageWidth - pDoc->m_margin[0] - pDoc->m_margin[2]);
		if (pDoc->m_columnWidth != columnWidth)
		{
			pDoc->m_columnWidth = columnWidth;
			pDoc->rpComputeDrawPositions();
			if (pDoc->m_ptrTableList != NULL)
			{
				pDoc->SetModifiedFlag(TRUE);
			}
		}
	}
	sprintf(s, "%d", pDoc->m_columnWidth);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	return;
}


void CMainFrame::OnReturn()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[32];

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);

	if (pDoc->m_return)
	{
		pDoc->m_return = FALSE;
		((CButton *)pCBox)->SetCheck(FALSE);

		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
		if (((CButton *)pCBox)->GetCheck())
		{
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_HIDE);
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_HIDE);
		}
	}
	else
	{
		pDoc->m_return = TRUE;
		((CButton *)pCBox)->SetCheck(TRUE);

		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
		if (((CButton *)pCBox)->GetCheck())
		{
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
			pCBox->ShowWindow(SW_SHOW);
			sprintf(s, "%d", pDoc->m_returnSizeEnvelope);
			pCBox->SetWindowText(s);
			pCBox->UpdateWindow();
			pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
			pCBox->ShowWindow(SW_SHOW);
		}
	}

	pDoc->rpComputeDrawPositions();
	
	return;
}


void CMainFrame::OnLineSpacingDefaults()
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[32];

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_lineSpacingEnvelope = 16;
		pCBox = (CComboBox*)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", pDoc->m_lineSpacingEnvelope);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_lineSpacingList = 14;
		pCBox = (CComboBox*)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", pDoc->m_lineSpacingList);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_lineSpacing = 14;
		pCBox = (CComboBox*)m_wndDlgBarTop.GetDlgItem(IDC_LINE_SPACING);
		sprintf(s, "%d", pDoc->m_lineSpacing);
		pCBox->SetWindowText(s);
		pCBox->UpdateWindow();
	}

	pDoc->rpComputeDrawPositions();
	if (pDoc->m_ptrTableList != NULL)
	{
		pDoc->SetModifiedFlag(TRUE);
	}
	
	return;
}


void CMainFrame::OnColumnWidthDefaults()

{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CWnd *pCBox;
	char s[32];

	pDoc->m_columnWidth = 240;
	pCBox = (CComboBox*)m_wndDlgBarTop.GetDlgItem(IDC_COLUMN_WIDTH);
	sprintf(s, "%d", pDoc->m_columnWidth);
	pCBox->SetWindowText(s);
	pCBox->UpdateWindow();

	pDoc->rpComputeDrawPositions();
	if (pDoc->m_ptrTableList != NULL)
	{
		pDoc->SetModifiedFlag(TRUE);
	}
	
	return;
}


void CMainFrame::OnUpdateListUI(CCmdUI *pCmdUI)
{

	return;
}


void CMainFrame::OnUpdateBookUI(CCmdUI *pCmdUI)
{

	return;
}


void CMainFrame::OnUpdateEnvelopeUI(CCmdUI *pCmdUI)
{

	return;
}


void CMainFrame::OnUpdateLabelUI(CCmdUI *pCmdUI)
{

	return;
}


void CMainFrame::OnViewReturnAddress() 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressReturnAddressDlg returnAddress;
	int i;
	char s[256];
	char *t;
	CWnd *pCBox;

	// TODO: Add your control notification handler code here

	i = 0;
	for (t = pDoc->m_returnAddress; *t == ' '; t++);  // remove leading white space if any
	for (; *t != '\0'; t++)
	{
		if (*t == ';')
		{
			// cannot use \n for newlines in edit control
			// rather CR LF character combination (0x0D 0x0A)
			s[i] = 0x0D;
			s[i + 1] = 0x0A;
			i++;

			t++;
			while(*t == ' ' && *(t + 1) == ' ')
			{
				t++;
			}
		}
		else
		{
			s[i] = *t;
		}
		i++;
	}
	s[i] = '\0';

	returnAddress.m_returnAddress.Format("%s", (LPCTSTR)s);


	if (returnAddress.DoModal() == IDOK)
	{
		pDoc->SetModifiedFlag(TRUE);
		sprintf(s, "%s", (LPCTSTR)returnAddress.m_returnAddress);

		i = 0;
		for (t = s; *t != '\0'; t++)
		{
			if (*t == 0x0D && *(t + 1) == 0x0A)
			{
				pDoc->m_returnAddress[i] = ';';
				pDoc->m_returnAddress[i + 1] = ' ';
				*(t + 1) = ' ';
				i++;

				t++;
				while(*t == ' ' && *(t + 1) == ' ')
				{
					t++;
				}
			}
			else
			{
				pDoc->m_returnAddress[i] = *t;
			}
			i++;
		}
		pDoc->m_returnAddress[i] = '\0';

		// remove any trailing whitespace or semi-colins
		for (t = &pDoc->m_returnAddress[i]; *t == ';' || *t == ' ' || *t == '\0'; t--)
		{
			*t = '\0';
		}

		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
		if (((CButton *)pCBox)->GetCheck())
		{
			if (pDoc->m_returnAddress[0] == '\0')
			{
				pDoc->m_return = FALSE;
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
				pCBox->ShowWindow(SW_HIDE);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
				pCBox->ShowWindow(SW_HIDE);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
				pCBox->ShowWindow(SW_HIDE);
			}
			else
			{
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
				pCBox->ShowWindow(SW_SHOW);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
				pCBox->ShowWindow(SW_HIDE);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
				pCBox->ShowWindow(SW_HIDE);
			}
			pDoc->rpComputeDrawPositions();
		}

		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
		if (((CButton *)pCBox)->GetCheck())
		{
			if (pDoc->m_returnAddress[0] == '\0')
			{
				pDoc->m_return = FALSE;
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
				pCBox->ShowWindow(SW_HIDE);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
				pCBox->ShowWindow(SW_HIDE);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
				pCBox->ShowWindow(SW_HIDE);
			}
			else
			{
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN);
				pCBox->ShowWindow(SW_SHOW);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_SIZE);
				pCBox->ShowWindow(SW_SHOW);
				pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_RETURN_PERCENT);
				pCBox->ShowWindow(SW_SHOW);
			}
			pDoc->rpComputeDrawPositions();
		}

	}


	return;	
}


void CMainFrame::OnFontChoose() 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CHOOSEFONT cf; 
    LOGFONT lf; 
    HFONT hfont;
	CWnd *pCBox;

	cf.lStructSize = sizeof(CHOOSEFONT); 
    cf.hwndOwner = (HWND)NULL; 
    cf.hDC = (HDC)NULL;

	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_fontBook.GetLogFont(&lf);  // put current font information into choose font dialog
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_fontList.GetLogFont(&lf);  // put current font information into choose font dialog
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_fontLabel.GetLogFont(&lf);  // put current font information into choose font dialog
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		pDoc->m_fontEnvelope.GetLogFont(&lf);  // put current font information into choose font dialog
	}
    cf.lpLogFont = &lf; 
    cf.iPointSize = 0; 
    cf.Flags = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT; 
    cf.rgbColors = RGB(0,0,0); 
    cf.lCustData = 0L; 
    cf.lpfnHook = (LPCFHOOKPROC)NULL; 
    cf.lpTemplateName = (LPSTR)NULL; 
    cf.hInstance = (HINSTANCE) NULL; 
    cf.lpszStyle = (LPSTR)NULL; 
    cf.nFontType = SCREEN_FONTTYPE; 
    cf.nSizeMin = 0; 
    cf.nSizeMax = 0;

	// Display the CHOOSEFONT common-dialog box. 
    if (ChooseFont(&cf) == IDOK)
	{
		// Create a logical font based on the user's selection 
		hfont = CreateFontIndirect(cf.lpLogFont);

		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
		if (((CButton *)pCBox)->GetCheck())
		{
			pDoc->m_fontBook.m_hObject = hfont;
		}
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
		if (((CButton *)pCBox)->GetCheck())
		{
			pDoc->m_fontList.m_hObject = hfont;
		}
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
		if (((CButton *)pCBox)->GetCheck())
		{
			pDoc->m_fontLabel.m_hObject = hfont;
		}
		pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
		if (((CButton *)pCBox)->GetCheck())
		{
			pDoc->m_fontEnvelope.m_hObject = hfont;
		}
		pDoc->rpComputeDrawPositions();
		if (pDoc->m_ptrTableList != NULL)
		{
			pDoc->SetModifiedFlag(TRUE);
		}
	}
	
	return;	
}

void CMainFrame::OnFontRestoreDefault() 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CFont font;
	LOGFONT lf;
	HFONT hfont;
	CWnd *pCBox;

	
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		font.CreatePointFont(90, "Times New Roman", NULL);
		font.GetLogFont(&lf);
		hfont = CreateFontIndirect(&lf); 
		pDoc->m_fontBook.m_hObject = hfont;
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		font.CreatePointFont(90, "Times New Roman", NULL);
		font.GetLogFont(&lf);
		hfont = CreateFontIndirect(&lf); 
		pDoc->m_fontList.m_hObject = hfont;
	}
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		font.CreatePointFont(120, "Times New Roman", NULL);
		font.GetLogFont(&lf);
		hfont = CreateFontIndirect(&lf); 
		pDoc->m_fontLabel.m_hObject = hfont;
	}	
	pCBox = (CComboBox *)m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		font.CreatePointFont(120, "Times New Roman", NULL);
		font.GetLogFont(&lf);
		hfont = CreateFontIndirect(&lf); 
		pDoc->m_fontEnvelope.m_hObject = hfont;
	}

	pDoc->rpComputeDrawPositions();
	if (pDoc->m_ptrTableList != NULL)
	{
		pDoc->SetModifiedFlag(TRUE);
	}

	return;
}

void CMainFrame::OnGroupUpdate() 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroupDlg group;
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	CAddressGroup *dupGroupSelected = pDoc->m_groupSelected;

	if (group.DoModal() == IDOK)
	{
		//if (pDoc->m_groupSelected != NULL)
		//{
		//	sprintf(s, "%s", pDoc->m_groupSelected->m_groupName);
		//}

		// make duplicate group list the active group list and delete the current group list
		for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL;)
		{
			prevGroup = ptrGroup;
			ptrGroup = ptrGroup->m_next;
			delete prevGroup;
		}
		pDoc->m_ptrGroupList = group.m_ptrGroupListDup;

		pDoc->rpUpdateGroupMenu();
		pDoc->SetModifiedFlag(TRUE);

		//if (pDoc->m_groupSelected != NULL)
		//{
		//	for (ptrGroup = pDoc->m_ptrGroupList;
		//		ptrGroup != NULL && strcmp(s, ptrGroup->m_groupName) != 0;
		//		ptrGroup = ptrGroup->m_next);
		//	pDoc->m_groupSelected = ptrGroup;
		//}

	}
	else
	{
		//if (pDoc->m_groupSelected != NULL)
		//{
		//	sprintf(s, "%s", pDoc->m_groupSelected->m_groupName);
		//}

		// delete all nodes in the duplicate group list
		for (ptrGroup = group.m_ptrGroupListDup; ptrGroup != NULL;)
		{
			prevGroup = ptrGroup;
			ptrGroup = ptrGroup->m_next;
			delete prevGroup;
		}

		pDoc->m_groupSelected = dupGroupSelected;
		
	}

	return;		
}

void CMainFrame::OnUpdateGroupUpdate(CCmdUI* pCmdUI) 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	if (pDoc->m_ptrTableList == NULL)
	{
		pCmdUI->Enable(FALSE);
	}
	else
	{
		pCmdUI->Enable(TRUE);
	}

	return;
}

void CMainFrame::OnUpdateGroupOne(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[0]);

	return;	
}


void CMainFrame::OnUpdateGroupTwo(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[1]);

	return;			
}


void CMainFrame::OnUpdateGroupThree(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[2]);

	return;		
}


void CMainFrame::OnUpdateGroupFour(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[3]);

	return;		
}


void CMainFrame::OnUpdateGroupFive(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[4]);

	return;		
}


void CMainFrame::OnUpdateGroupSix(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[5]);

	return;		
}


void CMainFrame::OnUpdateGroupSeven(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[6]);

	return;		
}


void CMainFrame::OnUpdateGroupEight(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[7]);

	return;			
}


void CMainFrame::OnUpdateGroupNine(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[8]);

	return;				
}


void CMainFrame::OnUpdateGroupTen(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[9]);

	return;				
}


void CMainFrame::OnUpdateGroupEleven(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(TRUE);
	pCmdUI->SetCheck(m_allGroupMembersSelected[10]);

	return;				
}


void CMainFrame::OnUpdateGroupTwelve(CCmdUI* pCmdUI) 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	int k = 0;
	CAddressGroup *ptrGroup;

	if (pDoc->m_ptrGroupList == NULL)
	{
		pCmdUI->Enable(FALSE);
		pCmdUI->SetCheck(FALSE);
	}
	else
	{
		pCmdUI->Enable(TRUE);
		for (ptrGroup = pDoc->m_ptrGroupList; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
		{
			if (ptrGroup == pDoc->m_ptrGroupList
				|| (ptrGroup->m_prev != NULL
				&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
			{
				k++;
			}
		}
		pCmdUI->SetCheck(m_allGroupMembersSelected[k - 1]);
	}

	return;		
	
}



void CMainFrame::OnGroupOne() 
{
	rpSelectGroupMembers(1);
	return;		
}


void CMainFrame::OnGroupTwo() 
{
	rpSelectGroupMembers(2);
	return;		
}


void CMainFrame::OnGroupThree() 
{
	rpSelectGroupMembers(3);
	return;		
}


void CMainFrame::OnGroupFour() 
{
	rpSelectGroupMembers(4);
	return;			
}


void CMainFrame::OnGroupFive() 
{
	rpSelectGroupMembers(5);
	return;		
}


void CMainFrame::OnGroupSix() 
{
	rpSelectGroupMembers(6);
	return;		
}


void CMainFrame::OnGroupSeven() 
{
	rpSelectGroupMembers(7);
	return;		
}

void CMainFrame::OnGroupEight() 
{
	rpSelectGroupMembers(8);
	return;			
}


void CMainFrame::OnGroupNine() 
{
	rpSelectGroupMembers(9);
	return;		
}


void CMainFrame::OnGroupTen() 
{
	rpSelectGroupMembers(10);
	return;		
}


void CMainFrame::OnGroupEleven() 
{
	rpSelectGroupMembers(11);
	return;			
}


void CMainFrame::OnGroupTwelve() 
{
	rpSelectGroupMembers(12);
	return;			
}



void CMainFrame::OnViewLastNameFirst() 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	if (pDoc->m_lastNameFirst == TRUE)
	{
		pDoc->m_lastNameFirst = FALSE;
	}
	else
	{
		pDoc->m_lastNameFirst = TRUE;
	}

	pDoc->rpSortName();

	return;
}


void CMainFrame::OnUpdateViewLastNameFirst(CCmdUI* pCmdUI) 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	
	pCmdUI->SetCheck(pDoc->m_lastNameFirst);
	return;
}


void CMainFrame::OnEditCopy()
// called when user selects copy from the edit menu
// text in the current view window is copied to the clipboard
{
#define CLIPSIZE 205600  // maximum chars copied to clipboard

	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	char s[CLIPSIZE];
	char t[256];
	char *u;
	CAddressTable *ptrTable;
	CWnd *pCBox;
	int i;
	int j;
	BOOL drawLabel = FALSE;
	BOOL drawBook = FALSE;
	BOOL drawList = FALSE;
	BOOL drawEnvelope = FALSE;

	pCBox = m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawLabel = TRUE;
	}
	pCBox = m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawBook = TRUE;
	}
	pCBox = m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawList = TRUE;
	}
	pCBox = m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawEnvelope = TRUE;
	}

	i = 0;
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (ptrTable->m_selected
			&& ((drawLabel && ptrTable->m_numberOfLinesLabel > 1) || drawBook || drawEnvelope || drawList))
		{
			// copy name
			if (i != 0 && i < CLIPSIZE - 2)
			{
				s[i] = '\n';
				i++;
			}

			if (pDoc->m_lastNameFirst == FALSE)
			{
				sprintf(t, "%s", ptrTable->m_name);
			}
			else
			{
				sprintf(t, "%s", ptrTable->m_nameLastFirst);
			}
			u = t;
			for (; i < CLIPSIZE - 2 && *u != '\0'; i++)
			{
				s[i] = *u;
				u++;
			}
			if (!drawList)
			{
				s[i] = '\n';
				i++;
			}
			if (i >= CLIPSIZE - 1)
			{
				s[CLIPSIZE - 1] = '\0';
				rpCopyToClipboard(s);
				return;
			}

			// copy address
			for (j = 0; j < 9 && ptrTable->m_address[j][0] != '\0'; j++)
			{
				if (drawList && i < CLIPSIZE - 4)
				{
					s[i] = ' ';
					s[i + 1] = '*';
					s[i + 2] = ' ';
					i += 3;
				}

				sprintf(t, ptrTable->m_address[j]);
				u = t;
				for (; i < CLIPSIZE - 2 && *u != '\0'; i++)
				{
					s[i] = *u;
					if (!drawList && s[i] == ';')
					{
						s[i] = '\n';
						u++;  // skip white space after ;
					}
					u++;
				}

				if (!drawList)
				{
					s[i] = '\n';
					i++;
				}
				if (i >= CLIPSIZE - 1)
				{
					s[CLIPSIZE - 1] = '\0';
					rpCopyToClipboard(s);
					return;
				}
				if (drawLabel || drawEnvelope)
				{
					j = 9;
				}
			}

			// copy email
			if (drawList || drawBook)
			{
				for (j = 0; j < 9 && ptrTable->m_email[j][0] != '\0'; j++)
				{
					if (drawList && i < CLIPSIZE - 4)
					{
						s[i] = ' ';
						s[i + 1] = '*';
						s[i + 2] = ' ';
						i += 3;
					}

					if (ptrTable->m_alias[j][0] != '\0')
					{
						sprintf(t, "%s (%s)", ptrTable->m_email[j], ptrTable->m_alias[j]);
					}
					else
					{
						sprintf(t, "%s", ptrTable->m_email[j]);
					}

					u = t;
					for (; i < CLIPSIZE - 2 && *u != '\0'; i++)
					{
						s[i] = *u;
						if (!drawList && s[i] == ';')
						{
							s[i] = '\n';
							u++;  // skip white space after ;
						}
						u++;
					}

					if (!drawList)
					{
						s[i] = '\n';
						i++;
					}
					if (i >= CLIPSIZE - 1)
					{
						s[CLIPSIZE - 1] = '\0';
						rpCopyToClipboard(s);
						return;
					}				

				}
			}

			// copy phone
			if (drawList || drawBook)
			{
				for (j = 0; j < 9 && ptrTable->m_phone[j][0] != '\0'; j++)
				{
					if (drawList && i < CLIPSIZE - 4)
					{
						s[i] = ' ';
						s[i + 1] = '*';
						s[i + 2] = ' ';
						i += 3;
					}

					sprintf(t, "%s", ptrTable->m_phone[j]);

					u = t;
					for (; i < CLIPSIZE - 2 && *u != '\0'; i++)
					{
						s[i] = *u;
						if (!drawList && s[i] == ';')
						{
							s[i] = '\n';
							u++;  // skip white space after ;
						}
						u++;
					}

					if (!drawList)
					{
						s[i] = '\n';
						i++;
					}
					if (i >= CLIPSIZE - 1)
					{
						s[CLIPSIZE - 1] = '\0';
						rpCopyToClipboard(s);
						return;
					}				

				}
			}

			// copy notes
			if (drawList || drawBook)
			{
				for (j = 0; j < 9 && ptrTable->m_note[j][0] != '\0'; j++)
				{
					if (drawList && i < CLIPSIZE - 4)
					{
						s[i] = ' ';
						s[i + 1] = '*';
						s[i + 2] = ' ';
						i += 3;
					}

					sprintf(t, "%s", ptrTable->m_note[j]);

					u = t;
					for (; i < CLIPSIZE - 2 && *u != '\0'; i++)
					{
						s[i] = *u;
						if (!drawList && s[i] == ';')
						{
							s[i] = '\n';
							u++;  // skip white space after ;
						}
						u++;
					}

					if (!drawList)
					{
						s[i] = '\n';
						i++;
					}
					if (i >= CLIPSIZE - 1)
					{
						s[CLIPSIZE - 1] = '\0';
						rpCopyToClipboard(s);
						return;
					}				

				}
			}

		}
	}
	s[i] = '\0';

	rpCopyToClipboard(s);

	return;
	
}

void CMainFrame::OnUpdateEditCopy(CCmdUI* pCmdUI) 
{
	CAddressView *pView = (CAddressView *)GetActiveView();
	if (!pView || !pView->IsKindOf(RUNTIME_CLASS(CAddressView)))
	{
		return;
	}
	CAddressDoc *pDoc = pView->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}

	CAddressTable *ptrTable;
	
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (ptrTable->m_selected)
		{
			pCmdUI->Enable(TRUE);
			return;
		}
	}

	pCmdUI->Enable(FALSE);
	return;
}
