// AddressGroupDlg.cpp : implementation file
//

#include "stdafx.h"
#include "address.h"
#include "AddressGroupDlg.h"
#include "addressDoc.h"
#include "addressView.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddressGroupDlg dialog


CAddressGroupDlg::CAddressGroupDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddressGroupDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddressGroupDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CAddressGroupDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddressGroupDlg)
	DDX_Control(pDX, IDC_MEMBER_NAME, m_memberName);
	DDX_Control(pDX, IDC_GROUP_NAME, m_groupName);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CAddressGroupDlg, CDialog)
	//{{AFX_MSG_MAP(CAddressGroupDlg)
	ON_EN_UPDATE(IDC_NEW_GROUP, OnUpdateNewGroup)
	ON_LBN_SELCHANGE(IDC_GROUP_NAME, OnSelchangeGroupName)
	ON_LBN_SELCHANGE(IDC_MEMBER_NAME, OnSelchangeMemberName)
	ON_BN_CLICKED(IDC_NEW, OnNew)
	ON_BN_CLICKED(IDC_ALL, OnAll)
	ON_BN_CLICKED(IDC_NONE, OnNone)
	ON_BN_CLICKED(IDC_DELETE, OnDelete)
	ON_COMMAND(ID_FILE_EXPORT_CSV, OnFileExportCsv)
	ON_LBN_KILLFOCUS(IDC_MEMBER_NAME, OnKillfocusMemberName)
	ON_LBN_SETFOCUS(IDC_MEMBER_NAME, OnSetfocusMemberName)
	ON_LBN_KILLFOCUS(IDC_GROUP_NAME, OnKillfocusGroupName)
	ON_LBN_SETFOCUS(IDC_GROUP_NAME, OnSetfocusGroupName)
	ON_EN_SETFOCUS(IDC_NEW_GROUP, OnSetfocusNewGroup)
	ON_EN_KILLFOCUS(IDC_NEW_GROUP, OnKillfocusNewGroup)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressGroupDlg message handlers


void CAddressGroupDlg::OnUpdateNewGroup() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function to send the EM_SETEVENTMASK message to the control
	// with the ENM_UPDATE flag ORed into the lParam mask.
	
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	char s[64];
	char t[64];
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_NEW_GROUP);
	pCBox->GetWindowText(s, 64);
	pCBox = (CComboBox*)GetDlgItem(IDC_NEW);

	if (strlen(s) == 0)
	{
		pCBox->EnableWindow(FALSE);
		return;
	}

	for (i = 0; i < m_groupName.GetCount(); i++)
	{
		m_groupName.GetText(i, t);
		if (strcmp(s, t) == 0)
		{
			pCBox->EnableWindow(FALSE);
			return;
		}
	}
	pCBox->EnableWindow(TRUE);
	return;
}

BOOL CAddressGroupDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	m_init = FALSE;

	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return FALSE;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup;
	CAddressTable *ptrTable = NULL;
	int i;
	int k;
	char s[128];
	CWnd *pCBox;

	m_groupName.ResetContent();

	if (pDoc->m_ptrGroupList != NULL && pDoc->m_groupSelected == NULL)
	{
		pDoc->m_groupSelected = pDoc->m_ptrGroupList;
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
			m_groupName.AddString(ptrGroup->m_groupName);
			if (ptrGroup == pDoc->m_groupSelected)
			{
				m_groupName.SetCurSel(k);	
			}
			k++;
		}
		if (strcmp(ptrGroup->m_groupName, pDoc->m_groupSelected->m_groupName) == 0)
		{
			for (; ptrTable != NULL && ptrTable != ptrGroup->m_ptrTable; ptrTable = ptrTable->m_next)
			{
				if (pDoc->m_lastNameFirst == FALSE)
				{
					m_memberName.AddString(ptrTable->m_name);
				}
				else
				{
					m_memberName.AddString(ptrTable->m_nameLastFirst);
				}
				i++;
			}
			if (ptrTable != NULL)
			{
				if (pDoc->m_lastNameFirst == FALSE)
				{
					m_memberName.AddString(ptrTable->m_name);
				}
				else
				{
					m_memberName.AddString(ptrTable->m_nameLastFirst);
				}
				m_memberName.SetSel(i, TRUE);
				i++;
				ptrTable = ptrTable->m_next;
			}
		}
	}

	for (; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (pDoc->m_lastNameFirst == FALSE)
		{
			m_memberName.AddString(ptrTable->m_name);
		}
		else
		{
			m_memberName.AddString(ptrTable->m_nameLastFirst);
		}
	}

	// duplicate the group list
	m_ptrGroupListDup = NULL;
	for (curGroup = pDoc->m_ptrGroupList; curGroup != NULL; curGroup = curGroup->m_next)
	{
		ptrGroup = new CAddressGroup;
		ptrGroup->m_ptrTable = curGroup->m_ptrTable;
		sprintf(ptrGroup->m_groupName, "%s", curGroup->m_groupName);
		ptrGroup->rpInsertGroup(&m_ptrGroupListDup, pDoc->m_lastNameFirst);
	}

	if (m_groupName.GetCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(FALSE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(FALSE);
		m_memberName.EnableWindow(FALSE);
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(FALSE);
	}
	else if (m_memberName.GetSelCount() == m_memberName.GetCount())
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		m_memberName.EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(TRUE);
	}
	else if (m_memberName.GetSelCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
		m_memberName.EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(TRUE);
	}
	else
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		m_memberName.EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(TRUE);
	}

	pCBox = (CComboBox*)GetDlgItem(IDC_NEW_GROUP);
	if (m_groupName.GetCount() >= MAXIMUM_NUMBER_GROUPS)
	{
		sprintf(s, "(group maximum reached)");
		pCBox->SetWindowText(s);
		pCBox->EnableWindow(FALSE);
	}
	else
	{
		sprintf(s, "");
		pCBox->SetWindowText(s);
		pCBox->EnableWindow(TRUE);
	}

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_TEXT);
	sprintf(s, "Group: (%d of %d)\n", m_groupName.GetCount(), MAXIMUM_NUMBER_GROUPS);
	pCBox->SetWindowText(s);

	if (pDoc->m_ptrGroupList == NULL)
	{
		sprintf(s, "Enter a name for a group then click the New button");
	}
	else if (m_groupName.GetCount() == 1)
	{
		sprintf(s, "Select members for %s, or create a new group",
			pDoc->m_groupSelected->m_groupName);
	}
	else if (m_groupName.GetCount() == MAXIMUM_NUMBER_GROUPS)
	{
		sprintf(s, "Select members for %s, or modify an existing group",
			pDoc->m_groupSelected->m_groupName);
	}
	else
	{
		sprintf(s, "Select members for %s, create a new group, or modify an existing group",
			pDoc->m_groupSelected->m_groupName);
	}
	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText(s);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


void CAddressGroupDlg::OnSelchangeGroupName() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	int k;
	char s[64];
	CWnd *pCBox;

	m_memberName.SetSel(-1, FALSE);
	m_groupName.GetText(m_groupName.GetCurSel(), s);

	k = 0;
	i = 0;
	//pDoc->m_groupSelected = NULL;
	ptrTable = pDoc->m_ptrTableList;
	for (ptrGroup = m_ptrGroupListDup; ptrGroup != NULL; ptrGroup = ptrGroup->m_next)
	{
		//if (ptrGroup == m_ptrGroupListDup
		//	|| (ptrGroup->m_prev != NULL
		//	&& strcmp(ptrGroup->m_prev->m_groupName, ptrGroup->m_groupName) != 0))
		//{
		//	if (k == m_groupName.GetCurSel())
		//	{
		//		pDoc->m_groupSelected = ptrGroup;
		//	}
		//	k++;
		//}


		//if (pDoc->m_groupSelected != NULL
		//	&& strcmp(pDoc->m_groupSelected->m_groupName, ptrGroup->m_groupName) == 0)
		if (strcmp(s, ptrGroup->m_groupName) == 0)
		{
			for (; ptrTable != NULL && ptrTable != ptrGroup->m_ptrTable; ptrTable = ptrTable->m_next)
			{
				i++;
			}
			if (ptrTable != NULL)
			{
				m_memberName.SetSel(i, TRUE);
				if (k == 0)
				{
					pDoc->m_groupSelected = ptrGroup;
				}
				k++;
			}
		}
	}

	if (m_memberName.GetSelCount() == m_memberName.GetCount())
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}
	else if (m_memberName.GetSelCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
	}
	else
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}

	return;
	
}

void CAddressGroupDlg::OnSelchangeMemberName() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	char s[64];
	CWnd *pCBox;

	m_groupName.GetText(m_groupName.GetCurSel(), s);

	ptrTable = pDoc->m_ptrTableList;
	for (i = 0; i < m_memberName.GetCaretIndex(); i++)
	{
		ptrTable = ptrTable->m_next;
	}

	if (m_memberName.GetSel(m_memberName.GetCaretIndex()))
	{
		ptrGroup = new CAddressGroup;
		ptrGroup->m_ptrTable = ptrTable;
		sprintf(ptrGroup->m_groupName, "%s", s);
		ptrGroup->rpInsertGroup(&m_ptrGroupListDup, pDoc->m_lastNameFirst);

		// have m_currentGroup point to the lead entry in this group
		for (curGroup = ptrGroup; curGroup->m_prev != NULL
			&& strcmp(curGroup->m_groupName, curGroup->m_prev->m_groupName) == 0;
			curGroup = curGroup->m_prev);
		pDoc->m_groupSelected = curGroup;
	}
	else
	{
		for (ptrGroup = m_ptrGroupListDup;
			ptrGroup != NULL
			&& (ptrGroup->m_ptrTable != ptrTable
			|| strcmp(s, ptrGroup->m_groupName) != 0);
			ptrGroup = ptrGroup->m_next);

		if (ptrGroup != NULL)
		{
			pDoc->m_groupSelected = ptrGroup->m_next;
			if (ptrGroup->m_next != NULL)
			{
				ptrGroup->m_next->m_prev = ptrGroup->m_prev;
			}
			if (ptrGroup->m_prev != NULL)
			{
				ptrGroup->m_prev->m_next = ptrGroup->m_next;
			}
			if (ptrGroup == m_ptrGroupListDup)
			{
				m_ptrGroupListDup = ptrGroup->m_next;
			}
			delete ptrGroup;
		}
	}

	if (m_memberName.GetSelCount() == m_memberName.GetCount())
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}
	else if (m_memberName.GetSelCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
	}
	else
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}

	return;
}

void CAddressGroupDlg::OnNew() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	char s[64];
	char t[64];
	char *u;
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_NEW_GROUP);
	pCBox->GetWindowText(s, 32);

	// remove white space
	i = 0;
	for (u = s; i < 32 && *u != '\0'; u++)
	{
		for (; *u == ' '; u++);
		s[i] = *u;
		i++;
	}
	s[i] = '\0';

	for (i = 0; i < m_groupName.GetCount(); i++)
	{
		m_groupName.GetText(i, t);
		if (strcmp(s, t) < 0)
		{
			break;
		}
	}
	m_groupName.InsertString(i, s);
	m_groupName.SetCurSel(i);
	m_memberName.SetSel(-1, FALSE);

	// initially select those members who are currently selected in main program
	i = 0;
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (ptrTable != NULL && ptrTable->m_selected)
		{
			m_memberName.SetSel(i, TRUE);

			ptrGroup = new CAddressGroup;
			ptrGroup->m_ptrTable = ptrTable;
			sprintf(ptrGroup->m_groupName, "%s", s);
			ptrGroup->rpInsertGroup(&m_ptrGroupListDup, pDoc->m_lastNameFirst);

			// have m_currentGroup point to the lead entry in this group
			for (curGroup = ptrGroup; curGroup->m_prev != NULL
				&& strcmp(curGroup->m_groupName, curGroup->m_prev->m_groupName) == 0;
				curGroup = curGroup->m_prev);
			pDoc->m_groupSelected = curGroup;
		}
		i++;
	}

	if (m_groupName.GetCount() >= MAXIMUM_NUMBER_GROUPS)
	{
		sprintf(s, "(group maximum reached)");
		pCBox->SetWindowText(s);
		pCBox->EnableWindow(FALSE);
	}
	else
	{
		sprintf(s, "");
		pCBox->SetWindowText(s);
		pCBox->EnableWindow(TRUE);
	}

	pCBox = (CComboBox*)GetDlgItem(IDC_NEW);
	pCBox->EnableWindow(FALSE);

	//pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
	//((CButton *)pCBox)->SetCheck(FALSE);
	//pCBox->EnableWindow(TRUE);

	//pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
	//((CButton *)pCBox)->SetCheck(TRUE);
	//pCBox->EnableWindow(TRUE);
	if (m_memberName.GetSelCount() == m_memberName.GetCount())
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}
	else if (m_memberName.GetSelCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(TRUE);
		pCBox->EnableWindow(TRUE);
	}
	else
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_ALL);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
		pCBox = (CComboBox*)GetDlgItem(IDC_NONE);
		((CButton *)pCBox)->SetCheck(FALSE);
		pCBox->EnableWindow(TRUE);
	}




	if (m_groupName.GetCount() == 1)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(TRUE);
	}

	m_memberName.EnableWindow(TRUE);

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_TEXT);
	sprintf(s, "Group: (%d of %d)\n", m_groupName.GetCount(), MAXIMUM_NUMBER_GROUPS);
	pCBox->SetWindowText(s);

	return;	
}

void CAddressGroupDlg::OnAll() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *curGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	CAddressTable *ptrTable = NULL;
	int i;
	char s[64];

	m_groupName.GetText(m_groupName.GetCurSel(), s);

	// first remove any group members from m_ptrGroupListDup
	for (ptrGroup = m_ptrGroupListDup; ptrGroup != NULL;)
	{
		if (ptrGroup != NULL && strcmp(s, ptrGroup->m_groupName) == 0)
		{
			if (ptrGroup->m_next != NULL)
			{
				ptrGroup->m_next->m_prev = ptrGroup->m_prev;
			}
			if (ptrGroup->m_prev != NULL)
			{
				ptrGroup->m_prev->m_next = ptrGroup->m_next;
			}
			if (ptrGroup == m_ptrGroupListDup)
			{
				m_ptrGroupListDup = ptrGroup->m_next;
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

	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		ptrGroup = new CAddressGroup;
		ptrGroup->m_ptrTable = ptrTable;
		sprintf(ptrGroup->m_groupName, "%s", s);
		if (ptrTable == pDoc->m_ptrTableList)
		{
			curGroup = ptrGroup;
			ptrGroup->rpInsertGroup(&m_ptrGroupListDup, pDoc->m_lastNameFirst);
			pDoc->m_groupSelected = ptrGroup;
		}
		else
		{
			ptrGroup->m_prev = curGroup;
			ptrGroup->m_next = curGroup->m_next;
			curGroup->m_next = ptrGroup;
			curGroup = ptrGroup;
		}
		
	}

	i = m_memberName.GetCount() - 1;
	if (i > 0)
	{
		m_memberName.SelItemRange(TRUE, 0, i);
	}
	else
	{
		m_memberName.SetSel(0, TRUE);
	}

	return;
}

void CAddressGroupDlg::OnNone() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	CAddressTable *ptrTable = NULL;
	char s[64];

	m_groupName.GetText(m_groupName.GetCurSel(), s);

	for (ptrGroup = m_ptrGroupListDup; ptrGroup != NULL;)
	{
		if (ptrGroup != NULL && strcmp(s, ptrGroup->m_groupName) == 0)
		{
			pDoc->m_groupSelected = ptrGroup->m_next;
			if (ptrGroup->m_next != NULL)
			{
				ptrGroup->m_next->m_prev = ptrGroup->m_prev;
			}
			if (ptrGroup->m_prev != NULL)
			{
				ptrGroup->m_prev->m_next = ptrGroup->m_next;
			}
			if (ptrGroup == m_ptrGroupListDup)
			{
				m_ptrGroupListDup = ptrGroup->m_next;
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
		
	m_memberName.SelItemRange(FALSE, 0, m_memberName.GetCount() - 1);

	return;
}

void CAddressGroupDlg::OnDelete() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}
	CView *pView = pMainFrame->GetActiveView();
	CAddressDoc *pDoc = ((CAddressView *)pView)->GetDocument();
	if (!pDoc || !pDoc->IsKindOf(RUNTIME_CLASS(CAddressDoc)))
	{
		return;
	}
	CAddressGroup *ptrGroup = NULL;
	CAddressGroup *prevGroup = NULL;
	CAddressTable *ptrTable = NULL;
	char s[64];
	CWnd *pCBox;

	m_groupName.GetText(m_groupName.GetCurSel(), s);
	m_groupName.DeleteString(m_groupName.GetCurSel());

	for (ptrGroup = m_ptrGroupListDup; ptrGroup != NULL;)
	{
		if (ptrGroup != NULL && strcmp(s, ptrGroup->m_groupName) == 0)
		{
			pDoc->m_groupSelected = ptrGroup->m_next;
			if (ptrGroup->m_next != NULL)
			{
				ptrGroup->m_next->m_prev = ptrGroup->m_prev;
			}
			if (ptrGroup->m_prev != NULL)
			{
				ptrGroup->m_prev->m_next = ptrGroup->m_next;
			}
			if (ptrGroup == m_ptrGroupListDup)
			{
				m_ptrGroupListDup = ptrGroup->m_next;
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

	if (m_groupName.GetCount() == 0)
	{
		pCBox = (CComboBox*)GetDlgItem(IDC_DELETE);
		pCBox->EnableWindow(FALSE);
		m_memberName.SelItemRange(FALSE, 0, m_memberName.GetCount() - 1);
		m_memberName.EnableWindow(FALSE);
	}
	else
	{
		m_groupName.SetCurSel(0);
		OnSelchangeGroupName();
	}

	pCBox = (CComboBox*)GetDlgItem(IDC_NEW_GROUP);
	sprintf(s, "");
	pCBox->SetWindowText(s);
	pCBox->EnableWindow(TRUE);

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_TEXT);
	sprintf(s, "Group: (%d of %d)\n", m_groupName.GetCount(), MAXIMUM_NUMBER_GROUPS);
	pCBox->SetWindowText(s);

	return;	
}

void CAddressGroupDlg::OnFileExportCsv() 
{
	// TODO: Add your command handler code here
	
}

void CAddressGroupDlg::OnKillfocusMemberName() 
{
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText("");
	return;
}

void CAddressGroupDlg::OnSetfocusMemberName() 
{
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText("");
	return;
	
}

void CAddressGroupDlg::OnKillfocusGroupName() 
{
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText("");
	return;	
}

void CAddressGroupDlg::OnSetfocusGroupName() 
{
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText("");
	return;	
}

void CAddressGroupDlg::OnSetfocusNewGroup() 
{
	CWnd *pCBox;

	if (m_init == FALSE)
	{
		m_init = TRUE;
		return;
	}

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);

	if (m_groupName.GetCurSel() == LB_ERR)
	{
		pCBox->SetWindowText("Enter a name for a group then click the New button");
	}
	else
	{
		pCBox->SetWindowText("Enter a unique name, then click the New button");
	}
	return;		
}

void CAddressGroupDlg::OnKillfocusNewGroup() 
{
	CWnd *pCBox;

	pCBox = (CComboBox*)GetDlgItem(IDC_GROUP_MESSAGE);
	pCBox->SetWindowText("");
	return;	
}
