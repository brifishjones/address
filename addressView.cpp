// addressView.cpp : implementation of the CAddressView class
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

#include <afxpriv.h>

/////////////////////////////////////////////////////////////////////////////
// CAddressView

IMPLEMENT_DYNCREATE(CAddressView, CScrollView)

BEGIN_MESSAGE_MAP(CAddressView, CScrollView)
	//{{AFX_MSG_MAP(CAddressView)
	ON_WM_SIZE()
	ON_WM_LBUTTONUP()
	ON_WM_RBUTTONUP()
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, OnFilePrintPreview)
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, CScrollView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, CScrollView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, CScrollView::OnFilePrintPreview)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddressView construction/destruction

CAddressView::CAddressView()
{
	// TODO: add construction code here

}

CAddressView::~CAddressView()
{
}

BOOL CAddressView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CScrollView::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CAddressView drawing

void CAddressView::OnDraw(CDC* pDC)
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	// TODO: add draw code for native data here

	CAddressTable *ptrTable;
	CFont font;
	CPen pen;
	CPen *ptrOldPen;
	CFont *ptrOldFont;
	CFont *ptrOldFontEnv;
	CFont retFont;
	CRect curCellRect;
	int i;
	int j;
	int k;
	int m;
	int pages;
	int lineSpacing;
	int labelCount;
	int numberOfLines;
	int fontSize;
	int offsetX;
	int offsetY;
	LOGFONT lf;
	HFONT hfont;
	char s[4096];
	char *t;
	char *u;
	CWnd *pCBox;
	BOOL drawLabel = FALSE;
	BOOL drawBook = FALSE;
	BOOL drawList = FALSE;
	BOOL drawEnvelope = FALSE;

	COLORREF black = RGB(0, 0, 0);
	COLORREF gray128 = RGB(128, 128, 128);
	COLORREF slateGray = 0x708090;
	COLORREF darkSlateGray = 0x2f4f4f;
	COLORREF slateBlue = 0x6a5acd;
	COLORREF darkSlateBlue = 0x483d8b;
	COLORREF teal = 0x008080;
	COLORREF darkOliveGreen = 0x556b2f;
	COLORREF darkCyan = 0x008b8b;

	//lineSpacing = pDoc->m_lineSpacing;

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LABEL);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawLabel = TRUE;
		labelCount = 0;
		//pDoc->m_lineSpacing = 16;
	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawBook = TRUE;
	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawList = TRUE;
	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_ENVELOPE);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawEnvelope = TRUE;
	}


	TRACE1("Printing = %d\n", pDC->IsPrinting());

	if (!pen.CreatePen(PS_SOLID, 1, gray128))
	{
		return;
	}

	if (drawBook)
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontBook);
	}
	if (drawList)
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontList);
	}
	if (drawLabel)
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontLabel);
	}
	if (drawEnvelope)
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontEnvelope);
		if (pDoc->m_return && pDoc->m_returnSizeEnvelope < 100)
		{
			pDoc->m_fontEnvelope.GetLogFont(&lf);
			lf.lfHeight = (LONG)((float)lf.lfHeight * (float)pDoc->m_returnSizeEnvelope * 0.01 + 0.5);
			hfont = CreateFontIndirect(&lf); 
			retFont.m_hObject = hfont;
		}

	}

	ptrOldPen = pDC->SelectObject(&pen);

	if (!pDC->IsPrinting() || pDoc->m_ptrTableList == NULL)
	{
		pages = pDoc->rpGetDocSize().cy / pDoc->m_printableHeight;
		for (i = 0; i < pages; i++)
		{
			if (!drawLabel)
			// draw edges of printable page
			{
				pDC->MoveTo(CPoint(0, i * -pDoc->m_printableHeight));
				pDC->LineTo(CPoint(pDoc->m_printableWidth, i * -pDoc->m_printableHeight));
				pDC->LineTo(CPoint(pDoc->m_printableWidth, (i + 1) * -pDoc->m_printableHeight));
				pDC->LineTo(CPoint(0, (i + 1) * -pDoc->m_printableHeight));
				pDC->LineTo(CPoint(0, i * -pDoc->m_printableHeight));
			}

			if (drawLabel && i == 0)
			// draw a first page of empty labels
			{
				// horizontal lines
				for (j = 0; j <= pDoc->m_labelDimension[1]; j++)
				{
					for (k = 0; k < pDoc->m_labelDimension[0]; k++)
					{
						pDC->MoveTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * k,
							-((pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY) + j * pDoc->m_labelHeight)));
						pDC->LineTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * k + pDoc->m_labelWidth,
							-((pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY) + j * pDoc->m_labelHeight)));
						//pDC->MoveTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * k, -((50 - pDoc->m_physicalOffsetY) + j * 100)));
						//pDC->LineTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * k + 262, -((50 - pDoc->m_physicalOffsetY) + j * 100)));
					}
				}
				// vertical lines
				for (j = 0; j < pDoc->m_labelDimension[0]; j++)
				{
					pDC->MoveTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * j,
						-(pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY)));
					pDC->LineTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * j,
						-((pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY) + pDoc->m_labelHeight * pDoc->m_labelDimension[1])));
					pDC->MoveTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * j + pDoc->m_labelWidth,
						-(pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY)));
					pDC->LineTo(CPoint(pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX + pDoc->m_labelColumnWidth * j + pDoc->m_labelWidth,
						-((pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY) +  pDoc->m_labelHeight * pDoc->m_labelDimension[1])));
					//pDC->MoveTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * j, -(50 - pDoc->m_physicalOffsetY)));
					//pDC->LineTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * j, -((50 - pDoc->m_physicalOffsetY) + 1000)));
					//pDC->MoveTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * j + 262, -(50 - pDoc->m_physicalOffsetY)));
					//pDC->LineTo(CPoint(12 - pDoc->m_physicalOffsetX + 283 * j + 262, -((50 - pDoc->m_physicalOffsetY) + 1000)));
				}

			}
		}
	}

	/**
	for (i = pDoc->m_offsetBorder; i <= pDoc->m_tableSize[1] + pDoc->m_offsetBorder; i += pDoc->m_cellSpacing[1])
	{
		pDC->MoveTo(CPoint(pDoc->m_offsetBorder, i));
		pDC->LineTo(CPoint(pDoc->m_tableSize[0] + pDoc->m_offsetBorder, i));
	}

	for (i = pDoc->m_offsetBorder; i <= pDoc->m_tableSize[0] + pDoc->m_offsetBorder; i += pDoc->m_cellSpacing[0])
	{
		pDC->MoveTo(CPoint(i, pDoc->m_offsetBorder));
		pDC->LineTo(CPoint(i, pDoc->m_tableSize[1] + pDoc->m_offsetBorder));
	}
	***/

	pDC->SetTextColor(black);
	pDC->SetBkMode(TRANSPARENT);


	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
	k = 0;

	// draw values for workcells and programs in table for current doc
	i = 0;	
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (((CListBox *)pCBox)->GetSel(k)
			&& ((drawLabel && ptrTable->m_numberOfLinesLabel > 1) || drawBook || drawEnvelope || drawList))
		{
			if (drawLabel)
			{
				labelCount++;
			}
			i = 0;
			if (drawBook)
			{
				pDC->MoveTo(CPoint(ptrTable->m_drawCoord[0], ptrTable->m_drawCoord[1] - 1));
				pDC->LineTo(CPoint(ptrTable->m_drawCoord[0] + pDoc->m_columnWidth, ptrTable->m_drawCoord[1] - 1));
			}

			if (drawLabel && !pDC->IsPrinting())
			{
				pDC->MoveTo(CPoint(ptrTable->m_drawCoord[0], ptrTable->m_drawCoord[1]));
				pDC->LineTo(CPoint(ptrTable->m_drawCoord[0] + pDoc->m_labelWidth, ptrTable->m_drawCoord[1]));
				pDC->LineTo(CPoint(ptrTable->m_drawCoord[0] + pDoc->m_labelWidth, ptrTable->m_drawCoord[1] - pDoc->m_labelHeight));
				pDC->LineTo(CPoint(ptrTable->m_drawCoord[0], ptrTable->m_drawCoord[1] - pDoc->m_labelHeight));
				pDC->LineTo(CPoint(ptrTable->m_drawCoord[0], ptrTable->m_drawCoord[1]));
			}


			if (ptrTable->m_name[0] != '\0')
			{

				if (drawEnvelope && pDoc->m_return && pDoc->m_returnAddress[0] != '\0')
				{
					if (pDoc->m_returnSizeEnvelope < 100)
					{
						ptrOldFontEnv = pDC->SelectObject(&retFont);
						lineSpacing = (int)((float)pDoc->m_lineSpacingEnvelope * (float)pDoc->m_returnSizeEnvelope * 0.01 + 0.5);
					}
					else
					{
						lineSpacing = pDoc->m_lineSpacingEnvelope;
					}

					u = pDoc->m_returnAddress;
					for (t = pDoc->m_returnAddress; *t != '\0'; t++)
					{
						if (*t == ';')
						{
							*t = '\0';
							//curCellRect.left = pDoc->m_physicalOffsetX;
							curCellRect.left = MAX(pDoc->m_physicalOffsetX, 15);  // some printers jam if too close to left edge
							curCellRect.top = ptrTable->m_drawCoord[1] + (pDoc->m_printableHeight / 2) - lineSpacing * i;
							curCellRect.right = pDoc->m_printableWidth;
							curCellRect.bottom = curCellRect.top - lineSpacing * 2;
							sprintf(s, "%s", u);
							//pDC->SetTextColor(darkSlateGray);
							pDC->DrawText(s, curCellRect, DT_LEFT | DT_NOPREFIX);
							i++;
							*t = ';';
							for (u = t + 1; *u == ' '; u++);
						}
					}
					//curCellRect.left = pDoc->m_physicalOffsetX;
					curCellRect.left = MAX(pDoc->m_physicalOffsetX, 15);  // some printers jam if too close to left edge
					curCellRect.top = ptrTable->m_drawCoord[1] + (pDoc->m_printableHeight / 2) - lineSpacing * i;
					curCellRect.right = pDoc->m_printableWidth;
					curCellRect.bottom = curCellRect.top - lineSpacing * 2;
					sprintf(s, "%s", u);
					//pDC->SetTextColor(darkSlateGray);
					pDC->DrawText(s, curCellRect, DT_LEFT | DT_NOPREFIX);
					i = 0;

					if (pDoc->m_returnSizeEnvelope < 100)
					{
						pDC->SelectObject(ptrOldFontEnv);
					}

				}

				if (drawBook)
				{
					curCellRect.left = ptrTable->m_drawCoord[0];
					curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
					curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
					curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
				}
				if (drawLabel)
				{
					pDoc->m_fontLabel.GetLogFont(&lf);
					fontSize = -MulDiv(lf.lfHeight, 72, 100);//pDC->GetDeviceCaps(LOGPIXELSY));
					if (ptrTable->m_numberOfLinesLabel <= 6)
					{
						lineSpacing = 3 * (11 - ptrTable->m_numberOfLinesLabel)
							* pDoc->m_labelHeight / 100;
					}
					else
					{
						lineSpacing = (pDoc->m_labelHeight - 3) / ptrTable->m_numberOfLinesLabel;

					}
					j = pDC->DrawText(s, curCellRect, DT_CALCRECT);
					curCellRect.left = ptrTable->m_drawCoord[0] + (pDoc->m_labelColumnWidth / 10) - 1;
					curCellRect.top = ptrTable->m_drawCoord[1]
						- (pDoc->m_labelHeight - (MIN(pDoc->m_labelHeight - 2, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
						- lineSpacing * i
						+ fontSize / 6 + 1;
					curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
					curCellRect.bottom = curCellRect.top - lineSpacing * 2;

					//pDoc->m_fontLabel.GetLogFont(&lf);
					//fontSize = -MulDiv(lf.lfHeight, 72, 100);//pDC->GetDeviceCaps(LOGPIXELSY));
					//if (ptrTable->m_numberOfLinesLabel <= 6)
					//{
					//	lineSpacing = 3 * (11 - ptrTable->m_numberOfLinesLabel);
					//}
					//else
					//{
					//	lineSpacing = 97 / ptrTable->m_numberOfLinesLabel;
					//}
					//j = pDC->DrawText(s, curCellRect, DT_CALCRECT);
					//curCellRect.left = ptrTable->m_drawCoord[0] + 25;
					//curCellRect.top = ptrTable->m_drawCoord[1]
					//	- (100 - (MIN(98, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
					//	- lineSpacing * i
					//	+ fontSize / 6 + 1;
					//curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
					//curCellRect.bottom = curCellRect.top - lineSpacing * 2;

					//TRACE3("fontSize = %d  lineSpacing = %d   offset = %d   ", fontSize, lineSpacing,
					//	-(100 - (MIN(96, lineSpacing * ptrTable->m_numberOfLinesLabel) - (lineSpacing - fontSize))) / 2
					//	- lineSpacing * i
					//	+ fontSize / 6 + 1);
					//TRACE1("height = %d\n", j);

					//lineSpacing = MAX(16, 96 / ptrTable->m_numberOfLinesLabel);
					//curCellRect.left = ptrTable->m_drawCoord[0] + 25;
					//curCellRect.top = ptrTable->m_drawCoord[1] - lineSpacing * i;
					//curCellRect.right = ptrTable->m_drawCoord[0] + 262;
					//curCellRect.bottom = curCellRect.top - lineSpacing * 2;
				}
				if (drawEnvelope)
				{
					curCellRect.left = ptrTable->m_drawCoord[0];
					curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingEnvelope * i;
					curCellRect.right = pDoc->m_printableWidth;
					curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingEnvelope * 2;
				}

				if (drawLabel || drawEnvelope)
				{
					// if last name first put at end of name string
					//sprintf(name, "%s", ptrTable->m_name);
					//u = name;
					//for (t = name; *t != '\0' && *t != ','; t++);
					//if (*t == ',')
					//{
					//	*t = '\0';
					//	t++;
					//	for (; *t == ' ' && *t != '\0'; t++);
					//	sprintf(s, "%s %s", t, u);
					//}
					//else
					//{
						sprintf(s, "%s", ptrTable->m_name);
					//}
				}
				else if (drawBook || drawList)
				{
					if (pDoc->m_lastNameFirst == FALSE)
					{
						sprintf(s, "%s", ptrTable->m_name);
					}
					else
					{
						sprintf(s, "%s", ptrTable->m_nameLastFirst);
					}
				}

				if (!drawList)
				{
					//pDC->SetTextColor(black);
					pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
					i++;
				}
			}

			if (drawList)
			{
				for (j = 0; j < 100 && ptrTable->m_listLineBreak[j] != NULL; j++)
				{
					if (j == 0)
					{
						sprintf(s, "%s", ptrTable->m_listLineBreak[j]);
					}
					else
					{
						sprintf(s, "     %s", ptrTable->m_listLineBreak[j]);
					}
					curCellRect.left = ptrTable->m_drawCoord[0];
					curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
					curCellRect.right = pDoc->m_printableWidth;
					curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
					pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
					i++;

				}

			}

			for (j = 0; j < 9 && ptrTable->m_address[j][0] != '\0'; j++)
			{
				if (drawList)
				{
					//if (ptrTable->m_address[j][255] == '\n')
					//{
					//	curCellRect.left = ptrTable->m_drawCoord[0];
					//	curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
					//	curCellRect.right = pDoc->m_printableWidth;
					//	curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
					//	pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
					//	i++;
					//	sprintf(s, "     %s", ptrTable->m_address[j]);
					//}
					//else
					//{
					//	sprintf(s, "%s * %s", s, ptrTable->m_address[j]);
					//}
				}
				else
				{
					u = ptrTable->m_address[j];
					for (t = ptrTable->m_address[j]; *t != '\0'; t++)
					{
						if (*t == ';')
						{
							*t = '\0';
							if (drawBook)
							{
								curCellRect.left = ptrTable->m_drawCoord[0];
								curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
								curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
								curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
							}
							if (drawLabel)
							{
								curCellRect.left = ptrTable->m_drawCoord[0] + (pDoc->m_labelColumnWidth / 10) - 1;
								curCellRect.top = ptrTable->m_drawCoord[1]
									- (pDoc->m_labelHeight - (MIN(pDoc->m_labelHeight - 2, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
									- lineSpacing * i
									+ fontSize / 6 + 1;
								curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
								curCellRect.bottom = curCellRect.top - lineSpacing * 2;

								//curCellRect.left = ptrTable->m_drawCoord[0] + 25;
								//curCellRect.top = ptrTable->m_drawCoord[1]
								//	- (100 - (MIN(96, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
								//	- lineSpacing * i
								//	+ fontSize / 6 + 1;
								//curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
								//curCellRect.bottom = curCellRect.top - lineSpacing * 2;


							}
							if (drawEnvelope)
							{
								curCellRect.left = ptrTable->m_drawCoord[0];
								curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingEnvelope * i;
								curCellRect.right = pDoc->m_printableWidth;
								curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingEnvelope * 2;
							}
							sprintf(s, "%s", u);
							//pDC->SetTextColor(darkSlateGray);
							pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
							i++;

							*t = ';';
							for (u = t + 1; *u == ' '; u++);
						}
					}
				}
				if (drawBook)
				{
					curCellRect.left = ptrTable->m_drawCoord[0];
					curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
					curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
					curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
				}
				if (drawLabel)
				{
					curCellRect.left = ptrTable->m_drawCoord[0] + (pDoc->m_labelColumnWidth / 10) - 1;
					curCellRect.top = ptrTable->m_drawCoord[1]
						- (pDoc->m_labelHeight - (MIN(pDoc->m_labelHeight - 2, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
						- lineSpacing * i
						+ fontSize / 6 + 1;
					curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
					curCellRect.bottom = curCellRect.top - lineSpacing * 2;

					//curCellRect.left = ptrTable->m_drawCoord[0] + 25;
					//curCellRect.top = ptrTable->m_drawCoord[1]
					//	- (100 - (MIN(96, lineSpacing * ptrTable->m_numberOfLinesLabel) - MAX(0, (lineSpacing - fontSize)))) / 2
					//	- lineSpacing * i
					//	+ fontSize / 6 + 1;
					//curCellRect.right = ptrTable->m_drawCoord[0] + pDoc->m_labelWidth;
					//curCellRect.bottom = curCellRect.top - lineSpacing * 2;
					j = 9; // draw only first address on label
				}
				if (drawEnvelope)
				{
					curCellRect.left = ptrTable->m_drawCoord[0];
					curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingEnvelope * i;
					curCellRect.right = pDoc->m_printableWidth;
					curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingEnvelope * 2;
					j = 9; // draw only first address on envelope
				}
				if (!drawList)
				{
					sprintf(s, "%s", u);
					//pDC->SetTextColor(darkSlateGray);
					pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
					i++;
				}
			}

			if (drawBook || drawList)
			{
				for (j = 0; j < 9 && ptrTable->m_email[j][0] != '\0'; j++)
				{
					if (drawList)
					{
						//if (ptrTable->m_email[j][255] == '\n')
						//{
						//	curCellRect.left = ptrTable->m_drawCoord[0];
						//	curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
						//	curCellRect.right = pDoc->m_printableWidth;
						//	curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
						//	pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
						//	i++;
						//	sprintf(s, "     %s", ptrTable->m_email[j]);
						//}
						//else
						//{
						//	sprintf(s, "%s * %s", s, ptrTable->m_email[j]);
						//}
					}
					else
					{
						curCellRect.left = ptrTable->m_drawCoord[0];
						curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
						curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
						curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
						//curCellRect.left = pDoc->m_offsetBorder;
						//curCellRect.top = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * i));
						//curCellRect.right = curCellRect.left + pDoc->m_cellSpacing[0] * 5;
						//curCellRect.bottom = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * (i + 1)));
						if (ptrTable->m_alias[j][0] != '\0')
						{
							sprintf(s, "%s (%s)", ptrTable->m_email[j], ptrTable->m_alias[j]);
						}
						else
						{
							sprintf(s, "%s", ptrTable->m_email[j]);
						}
						//pDC->SetTextColor(teal);
						pDC->DrawText(s, curCellRect, DT_LEFT);
						i++;
					}
				}
			}

			if (drawBook || drawList)
			{
				for (j = 0; j < 9 && ptrTable->m_phone[j][0] != '\0'; j++)
				{
					if (drawList)
					{
						//if (ptrTable->m_phone[j][255] == '\n')
						//{
						//	curCellRect.left = ptrTable->m_drawCoord[0];
						//	curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
						//	curCellRect.right = pDoc->m_printableWidth;
						//	curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
						//	pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
						//	i++;
						//	sprintf(s, "     %s", ptrTable->m_phone[j]);
						//}
						//else
						//{
						//	sprintf(s, "%s * %s", s, ptrTable->m_phone[j]);
						//}
					}
					else
					{
						curCellRect.left = ptrTable->m_drawCoord[0];
						curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
						curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
						curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
						//curCellRect.left = pDoc->m_offsetBorder;
						//curCellRect.top = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * i));
						//curCellRect.right = curCellRect.left + pDoc->m_cellSpacing[0] * 5;
						//curCellRect.bottom = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * (i + 1)));
						sprintf(s, "%s", ptrTable->m_phone[j]);
						//pDC->SetTextColor(darkSlateBlue);
						pDC->DrawText(s, curCellRect, DT_LEFT);
						i++;
					}
				}
			}

			if (drawBook || drawList)
			{
				for (j = 0; j < 9 && ptrTable->m_note[j][0] != '\0'; j++)
				{
					if (drawList)
					{
						//if (ptrTable->m_note[j][255] == '\n')
						//{
						//	curCellRect.left = ptrTable->m_drawCoord[0];
						//	curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
						//	curCellRect.right = pDoc->m_printableWidth;
						//	curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
						//	pDC->DrawText(s, curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
						//	i++;
						//	sprintf(s, "     %s", ptrTable->m_note[j]);
						//}
						//else
						//{
						//	sprintf(s, "%s * %s", s, ptrTable->m_note[j]);
						//}
					}
					else
					{
						curCellRect.left = ptrTable->m_drawCoord[0];
						curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacing * i;
						curCellRect.right = curCellRect.left + pDoc->m_columnWidth;
						curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacing * 2;
						//curCellRect.left = pDoc->m_offsetBorder;
						//curCellRect.top = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * i));
						//curCellRect.right = curCellRect.left + pDoc->m_cellSpacing[0] * 5;
						//curCellRect.bottom = -(pDoc->m_offsetBorder + (pDoc->m_lineSpacing * (i + 1)));
						sprintf(s, "%s", ptrTable->m_note[j]);
						//pDC->SetTextColor(darkSlateGray);
						pDC->DrawText(s, curCellRect, DT_LEFT | DT_NOPREFIX);
						i++;
					}
				}
			}

			if (drawList)
			{
				//curCellRect.left = ptrTable->m_drawCoord[0];
				//curCellRect.top = ptrTable->m_drawCoord[1] - pDoc->m_lineSpacingList * i;
				//curCellRect.right = pDoc->m_printableWidth;
				//curCellRect.bottom = curCellRect.top - pDoc->m_lineSpacingList * 2;
				//pDC->DrawText(s, curCellRect, DT_LEFT | DT_NOPREFIX);
			}

		}

		k++;
	}

	// draw return address labels
	if (drawLabel && pDoc->m_return)
	{
		u = pDoc->m_returnAddress;
		i = 0;
		for (t = pDoc->m_returnAddress; *t != '\0' && i < 6; t++)
		{
			if (*t == ';')
			{
				*t = '\0';
				sprintf(&s[i * 256], "%s", u);
				i++;
				*t = ';';
				for (u = t + 1; *u == ' '; u++);
			}
		}
		sprintf(&s[i * 256], "%s", u);

		numberOfLines = i + 1;
		//linesPerPage = 60;
		//columnsPerPage = 3;
		//columnOffset = pDoc->m_labelColumnWidth;
		pDoc->m_fontLabel.GetLogFont(&lf);
		fontSize = -MulDiv(lf.lfHeight, 72, 100);

		//lineSpacing = 3 * (11 - numberOfLines);
		if (numberOfLines <= 6)
		{
			lineSpacing = 3 * (11 - numberOfLines)
				* pDoc->m_labelHeight / 100;
		}
		else
		{
			lineSpacing = (pDoc->m_labelHeight - 3) / numberOfLines;
		}
		pDoc->m_lastLabel = MAX(pDoc->m_lastLabel, pDoc->m_firstLabel + labelCount);

		j = (pDoc->m_firstLabel + labelCount) / pDoc->m_labelDimension[1];   // column
		i = (pDoc->m_firstLabel + labelCount) % pDoc->m_labelDimension[1];   // row

		// get position (upper left corner) of starting return address label
		offsetY = ((j / pDoc->m_labelDimension[0]) * pDoc->m_printableHeight
			+ pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY)
			+ (((pDoc->m_firstLabel + labelCount) % pDoc->m_labelDimension[1]) * pDoc->m_labelHeight);

		offsetX = pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX
			+ (j * pDoc->m_labelColumnWidth) % (pDoc->m_labelDimension[0] * pDoc->m_labelColumnWidth);

		TRACE1("OnDraw() pDoc->m_lastLabel = %d\n", pDoc->m_lastLabel);

		for (k = 0; k <= pDoc->m_lastLabel - (pDoc->m_firstLabel + labelCount); k++)
		{
			i++;
			if (i > pDoc->m_labelDimension[1] && j % pDoc->m_labelDimension[0] == 2)
			{
				break;
			}
			else if (i > pDoc->m_labelDimension[1])
			{
				j++;
				i = 1;
				offsetY = (j / pDoc->m_labelDimension[0]) * pDoc->m_printableHeight
					+ pDoc->m_labelTopMargin - pDoc->m_physicalOffsetY;
				offsetX = pDoc->m_labelLeftMargin - pDoc->m_physicalOffsetX
					+ (j * pDoc->m_labelColumnWidth) % (pDoc->m_labelDimension[0] * pDoc->m_labelColumnWidth);
			}

			if (!pDC->IsPrinting())
			{
				pDC->MoveTo(CPoint(offsetX, -offsetY));
				pDC->LineTo(CPoint(offsetX + pDoc->m_labelWidth, -offsetY));
				pDC->LineTo(CPoint(offsetX + pDoc->m_labelWidth, -offsetY - pDoc->m_labelHeight));
				pDC->LineTo(CPoint(offsetX, -offsetY - pDoc->m_labelHeight));
				pDC->LineTo(CPoint(offsetX, -offsetY));
			}

			for (m = 0; m < numberOfLines; m++)
			{
				curCellRect.left = offsetX + (pDoc->m_labelColumnWidth / 10) - 1;
				curCellRect.top = -offsetY
					- (pDoc->m_labelHeight - (MIN(pDoc->m_labelHeight - 2, lineSpacing * numberOfLines) - MAX(0, (lineSpacing - fontSize)))) / 2
					- lineSpacing * m
					+ fontSize / 6 + 1;
				curCellRect.right = offsetX + pDoc->m_labelWidth;
				curCellRect.bottom = curCellRect.top - lineSpacing * 2;
				pDC->DrawText(&s[m * 256], curCellRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_NOPREFIX);
			}

			offsetY += pDoc->m_labelHeight;
		}

	}

	pDC->SelectObject(ptrOldFont);
	pDC->SelectObject(ptrOldPen);
	//pDoc->m_lineSpacing = lineSpacing;

	return;
}

/////////////////////////////////////////////////////////////////////////////
// CAddressView printing

BOOL CAddressView::OnPreparePrinting(CPrintInfo* pInfo)
{

	return DoPreparePrinting(pInfo);
}

void CAddressView::OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo)
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	pInfo->SetMaxPage(pDoc->rpComputeDrawPositions());

	TRACE0("OnBeginPrinting()\n");

	return;
	// TODO: add extra initialization before printing
}

void CAddressView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/////////////////////////////////////////////////////////////////////////////
// CAddressView diagnostics

#ifdef _DEBUG
void CAddressView::AssertValid() const
{
	CScrollView::AssertValid();
}

void CAddressView::Dump(CDumpContext& dc) const
{
	CScrollView::Dump(dc);
}

CAddressDoc* CAddressView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CAddressDoc)));
	return (CAddressDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CAddressView message handlers

void CAddressView::OnInitialUpdate() 
{
	CScrollView::OnInitialUpdate();
	
	// TODO: Add your specialized code here and/or call the base class
	//SetScrollSizes(MM_TEXT, GetDocument()->rpGetDocSize());
	SetScrollSizes(MM_LOENGLISH, GetDocument()->rpGetDocSize());
}

void CAddressView::OnUpdate(CView* pSender, LPARAM lHint, CObject* pHint) 
{
	// TODO: Add your specialized code here and/or call the base class
	//SetScrollSizes(MM_TEXT, GetDocument()->rpGetDocSize());
	SetScrollSizes(MM_LOENGLISH, GetDocument()->rpGetDocSize());
}

void CAddressView::OnPrint(CDC* pDC, CPrintInfo* pInfo) 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);


	// TODO: Add your specialized code here and/or call the base class

	//rpDrawBook(pDC, pInfo);

	//if (pInfo->m_nCurPage == 1)
	//{
	//	ptrTable = pDoc->m_ptrTableList;
	//}
	//else
	//{
	//	for (ptrTable = pDoc->m_ptrTableList;
	//	ptrTable != NULL && (ptrTable->m_columnBook - 1) / pDoc->m_columns < (int)pInfo->m_nCurPage;
	//	ptrTable = ptrTable->m_next);
	//}

	// If in print preview mode then the clipping
	// rectangle needs to be adjusted before creating the
	// clipping region

	CRgn pRgn;
	CRect rectClip;
	CPoint ptOrg;
	CPoint pageOrg;
	BOOL drawList = FALSE;
	BOOL drawBook = FALSE;
	CWnd *pCBox;

	rectClip.left = 0;
	rectClip.top = 0;
	//rectClip.right = rectClip.left + 400;
	//rectClip.bottom = -1000;

	pageOrg.x = 0;
	//pageOrg.y = -((pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]);
	//pageOrg.y = -((pInfo->m_nCurPage - 1) * 1047);
	pageOrg.y = (pInfo->m_nCurPage - 1) * -pDoc->m_printableHeight;


	//pDC->SetViewportOrg(0, -((pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]));

	/**
	if (pDC->IsKindOf(RUNTIME_CLASS(CPreviewDC)))
    {

		CPreviewDC *pPrevDC = (CPreviewDC *)pDC;

		pPrevDC->LPtoDP(&rectClip);
		pPrevDC->PrinterDPtoScreenDP(&rectClip.TopLeft());
		pPrevDC->PrinterDPtoScreenDP(&rectClip.BottomRight());

		// Now offset the result by the viewport origin of
		// the print preview window...

		::GetViewportOrgEx(pDC->m_hDC, &ptOrg);
		rectClip += ptOrg;

		//pPrevDC->LPtoDP(&pageOrg);
		//pPrevDC->PrinterDPtoScreenDP(&pageOrg);



	}
	**/
	
	pDC->SetWindowOrg(pageOrg);
//	pDC->SetViewportOrg(50, 25);

	// The following two function calls are the ones that
	// select the clipping region into the DC. These would be
	// whatever code you already have to create/select the
	// clipping region
	//pRgn.CreateRectRgn(rectClip.left, rectClip.top, rectClip.right, rectClip.bottom);
	//pDC->SelectClipRgn(&pRgn);


	TRACE2("OnPrint(): m_rectDraw[%d, %d] ", pInfo->m_rectDraw.left, pInfo->m_rectDraw.top);
	TRACE2("[%d, %d] ", pInfo->m_rectDraw.right, pInfo->m_rectDraw.bottom);
	TRACE2("H = %d  W = %d\n", pInfo->m_rectDraw.Height(), pInfo->m_rectDraw.Width());
	CPoint pt;
	CRect re;
	pt = pDC->GetWindowOrg();
	pDC->GetClipBox(&re);
	TRACE3("OnPrint(): page %d   window origin[%d, %d]  ",
		pInfo->m_nCurPage, pt.x, pt.y);
	TRACE2("Clip box[%d, %d], ", re.left, re.top);
	TRACE2("[%d, %d]\n", re.right, re.bottom);

	//ptOrg.x = 0;
	//ptOrg.y = -((pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]);
	//pDC->DPtoLP(&ptOrg);
	//pDC->SetViewportOrg(pageOrg);
	//pDC->SetViewportOrg(0, -((pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]));


	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawBook = TRUE;
	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		drawList = TRUE;
	}

	if (drawList || drawBook)
	{
		rpDrawPageHeader(pDC, pInfo);
	}


	CScrollView::OnPrint(pDC, pInfo);
}





void CAddressView::OnPrepareDC(CDC* pDC, CPrintInfo* pInfo) 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	// TODO: Add your specialized code here and/or call the base class


	if (pDC->IsPrinting())
	{

		CRgn pRgn;
		CRect rectClip;

		//rectClip.left = 0;
		//rectClip.top = 0;
		//rectClip.right = rectClip.left + 400;
		//rectClip.bottom = -800;

		//pDC->SetViewportOrg(0, -((pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]));

		//pDC->SetViewportExt(500, -500);
		//pDC->SetWindowOrg(pDoc->m_offsetBorder, -(pDoc->m_offsetBorder + (pInfo->m_nCurPage - 1) * pDoc->m_pageSize[1]));
		//pDC->SelectClipRgn(&pRgn);
		//pInfo->m_rectDraw.right = 600;
		//pInfo->m_rectDraw.bottom = -800;

		// If in print preview mode then the clipping
		// rectangle needs to be adjusted before creating the
		// clipping region
		//if (pDC->IsKindOf(RUNTIME_CLASS(CPreviewDC)))
        //{
		//CPreviewDC *pPrevDC = (CPreviewDC *)pDC;

		//	pPrevDC->PrinterDPtoScreenDP(&rectClip.TopLeft());
		//	pPrevDC->PrinterDPtoScreenDP(&rectClip.BottomRight());

			// Now offset the result by the viewport origin of
			// the print preview window...

		//	CPoint ptOrg;
		//	 ::GetViewportOrgEx(pDC->m_hDC,&ptOrg);
		//	rectClip += ptOrg;

		// The following two function calls are the ones that
		// select the clipping region into the DC. These would be
		// whatever code you already have to create/select the
		// clipping region

		//pRgn.CreateRectRgn(rectClip.left, rectClip.top, rectClip.right, rectClip.bottom);
		//pDC->SelectClipRgn(&pRgn);

	}
	
	TRACE2("OnPrepareDC(%d, %d) ", pDC, pInfo);
	TRACE2("LOGPIXELS(%d, %d) ", pDC->GetDeviceCaps(LOGPIXELSX), pDC->GetDeviceCaps(LOGPIXELSY));
	TRACE2("   HORZRES = %d   VERTRES = %d    ",
		pDC->GetDeviceCaps(HORZRES), // * 100 / pDC->GetDeviceCaps(LOGPIXELSX),
		pDC->GetDeviceCaps(VERTRES)); //* 100 / pDC->GetDeviceCaps(LOGPIXELSY));
	TRACE2("PHYSICALOFFSETX = %d   PHYSICALOFFSETY = %d\n",
		pDC->GetDeviceCaps(PHYSICALOFFSETX), // * 100 / pDC->GetDeviceCaps(LOGPIXELSX),
		pDC->GetDeviceCaps(PHYSICALOFFSETY)); // * 100 / pDC->GetDeviceCaps(LOGPIXELSY));

	CScrollView::OnPrepareDC(pDC, pInfo);
}

/**
int CAddressView::rpComputeDrawPositions()
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);


	CAddressTable *ptrTable;
	int linesPerPage;
	int columnsPerPage;
	int columnOffset;
	int offsetX;
	int offsetY;
	int i;
	int j;
	CDC dc;
	CPrintDialog dlg(FALSE);

	dlg.GetDefaults();
	dc.Attach(dlg.GetPrinterDC());

	pDoc->m_margin[0] = MAX(pDoc->m_margin[0],
		dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX));
	pDoc->m_margin[1] = MAX(pDoc->m_margin[1],
		dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	pDoc->m_margin[2] = MAX(pDoc->m_margin[2],
		dc.GetDeviceCaps(PHYSICALWIDTH) * 100 / dc.GetDeviceCaps(LOGPIXELSX)
		- dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX)
		- dc.GetDeviceCaps(HORZRES) * 100 / dc.GetDeviceCaps(LOGPIXELSX));
	pDoc->m_margin[3] = MAX(pDoc->m_margin[3],
		dc.GetDeviceCaps(PHYSICALHEIGHT) * 100 / dc.GetDeviceCaps(LOGPIXELSY)
		- dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY)
		- dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));

	TRACE2("rpComputeDrawPositions():  HORZRES = %d   VERTRES = %d\n",
		dc.GetDeviceCaps(HORZRES) * 100 / dc.GetDeviceCaps(LOGPIXELSX),
		dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY));
	TRACE2("PHYSICALOFFSETX = %d   PHYSICALOFFSETY = %d\n",
		dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX),
		dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY));


	linesPerPage
		= (dc.GetDeviceCaps(PHYSICALHEIGHT) * 100 / dc.GetDeviceCaps(LOGPIXELSY)
		- pDoc->m_margin[1]
		- pDoc->m_margin[3])
		/ pDoc->m_lineSpacing;
	columnsPerPage
		= (dc.GetDeviceCaps(PHYSICALWIDTH) * 100 / dc.GetDeviceCaps(LOGPIXELSX)
		- pDoc->m_margin[0]
		- pDoc->m_margin[2])
		/ pDoc->m_columnWidth;
	columnOffset = (dc.GetDeviceCaps(PHYSICALWIDTH) * 100 / dc.GetDeviceCaps(LOGPIXELSX)
		- pDoc->m_margin[0]
		- pDoc->m_margin[2])
		/ columnsPerPage;

	j = 0;
	i = 0;
	offsetX = pDoc->m_margin[0] - dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX);
	offsetY = pDoc->m_margin[1] - dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY);
	
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		i += ptrTable->m_numberOfLinesBook;
		if (i > linesPerPage)
		{
			j++;
			i = ptrTable->m_numberOfLinesBook;
			offsetY = (j / columnsPerPage) * dc.GetDeviceCaps(VERTRES) * 100 / dc.GetDeviceCaps(LOGPIXELSY)
				+ pDoc->m_margin[1] - dc.GetDeviceCaps(PHYSICALOFFSETY) * 100 / dc.GetDeviceCaps(LOGPIXELSY);
			offsetX = pDoc->m_margin[0] - dc.GetDeviceCaps(PHYSICALOFFSETX) * 100 / dc.GetDeviceCaps(LOGPIXELSX)
				+ (j * columnOffset) % (columnsPerPage * columnOffset);
			ptrTable->m_drawCoord[0] = offsetX;
			ptrTable->m_drawCoord[1] = -offsetY;
			offsetY += ptrTable->m_numberOfLinesBook * pDoc->m_lineSpacing;			
		}
		else
		{
			ptrTable->m_drawCoord[0] = offsetX;
			ptrTable->m_drawCoord[1] = -offsetY;
			offsetY += ptrTable->m_numberOfLinesBook * pDoc->m_lineSpacing;
		}
		ptrTable->m_columnBook = j;

	}

	return j;
}
**/

void CAddressView::rpDrawPageHeader(CDC *pDC, CPrintInfo *pInfo)
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	CPoint pageOrg;
	CRect curCellRect;
	CFont font;
	CPen pen;
	CFont *ptrOldFont;
	CWnd *pCBox;
	int lineSpacing;
	char s[256];

	pageOrg.x = 0;
	pageOrg.y = (pInfo->m_nCurPage - 1) * -pDoc->m_printableHeight;

	//if (!pen.CreatePen(PS_SOLID, 1, RGB(200, 200, 200)))
	//{
	//	return;
	//}

	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_BOOK);
	if (((CButton *)pCBox)->GetCheck())
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontBook);
		lineSpacing = pDoc->m_lineSpacing;
	}
	pCBox = (CComboBox*)pMainFrame->m_wndDlgBarTop.GetDlgItem(IDC_LIST);
	if (((CButton *)pCBox)->GetCheck())
	{
		ptrOldFont = pDC->SelectObject(&pDoc->m_fontList);
		lineSpacing = pDoc->m_lineSpacingList;
	}

	if (pageOrg.y - (pDoc->m_margin[1] - pDoc->m_physicalOffsetY) + lineSpacing <= pageOrg.y)
	{
		//ptrOldPen = pDC->SelectObject(&pen);
		pDC->FillSolidRect(pDoc->m_margin[0] - pDoc->m_physicalOffsetX,
			pageOrg.y - (pDoc->m_margin[1] - pDoc->m_physicalOffsetY) + lineSpacing,
			pDoc->m_pageWidth - pDoc->m_margin[3] - pDoc->m_margin[0],
			-lineSpacing,
			RGB(225, 225, 225));

		pDC->SetTextColor(RGB(0, 0, 0));
		pDC->SetBkMode(TRANSPARENT);

		curCellRect.left = pDoc->m_margin[0] - pDoc->m_physicalOffsetX;
		curCellRect.top = pageOrg.y - (pDoc->m_margin[1] - pDoc->m_physicalOffsetY) + lineSpacing;
		curCellRect.right = pDoc->m_pageWidth - pDoc->m_margin[3] - pDoc->m_physicalOffsetX;
		curCellRect.bottom = curCellRect.top - lineSpacing * 2;
		sprintf(s, "%d", pInfo->m_nCurPage);
		pDC->DrawText(s, curCellRect, DT_RIGHT | DT_NOPREFIX);
		sprintf(s, "%s", pDoc->GetTitle());
		pDC->DrawText(s, curCellRect, DT_CENTER | DT_NOPREFIX);

		sprintf(s, "%s", pDoc->GetPathName());
		CFileStatus status;
		if (CFile::GetStatus(s, status))
		{
			sprintf(s, "%04d.%02d.%02d",
				status.m_mtime.GetYear(), status.m_mtime.GetMonth(), status.m_mtime.GetDay());
			pDC->DrawText(s, curCellRect, DT_LEFT | DT_NOPREFIX);
		}

		TRACE3("rpDrawPageHeader %d (%d, %d,", pInfo->m_nCurPage, curCellRect.left, curCellRect.top);
		TRACE2(" %d, %d)\n", curCellRect.right, curCellRect.bottom);
	}

	pDC->SelectObject(ptrOldFont);
	//pDC->SelectObject(ptrOldPen);

	return;
}



BOOL CAddressView::rpLeftDlgListbox()
// size the address name list box in the left dialog box
// to be the same height as the view window 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return FALSE;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	RECT dlRect;
	RECT nRect;
	RECT rect;
	RECT vRect;
	CWnd *pCBox;

	this->GetWindowRect(&vRect);

	if (!::GetWindowRect(pMainFrame->m_wndDlgBarLeft.m_hWnd, &dlRect))
	// if left dialog has not yet been created then return
	{
		return FALSE;	
	}

	pMainFrame->m_wndDlgBarLeft.GetWindowRect(&dlRect);

	pCBox = (CComboBox *)pMainFrame->m_wndDlgBarLeft.GetDlgItem(IDC_NAME);

	((CWnd *)pCBox)->GetWindowRect(&nRect);

	((CListBox *)pCBox)->SetHorizontalExtent((nRect.right - nRect.left) + 75);

	rect.left = nRect.left - dlRect.left;
	rect.top = nRect.top - dlRect.top;
	rect.right = rect.left + (nRect.right - nRect.left);
	rect.bottom = (vRect.bottom - vRect.top);

	((CWnd *)pCBox)->MoveWindow(&rect, TRUE);

	return TRUE;
}

void CAddressView::OnSize(UINT nType, int cx, int cy) 
{
	CScrollView::OnSize(nType, cx, cy);
	
	// TODO: Add your message handler code here
	if ( 0 >= cx || 0 >= cy )
	{
		return;
	}

	rpLeftDlgListbox();
	return;
}


void CAddressView::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	CPoint sp = GetScrollPosition();
	int offsetX;
	int offsetY;
	int i;
	CWnd *pCBox;
	CAddressTable *ptrTable;

	CDC *pDC;
	pDC = GetDC( );

	// CScrollView changes the viewport origin and mapping mode.
	// It's necessary to convert the point from device coordinates
	// to logical coordinates, such as are stored in the document.
	CClientDC dc(this);
	OnPrepareDC(&dc);
	dc.DPtoLP(&point);

	offsetX = pDoc->m_physicalOffsetX + point.x;
	offsetY = pDoc->m_physicalOffsetY - point.y;// - pDoc->m_labelTopMargin;
	// TODO: Add your message handler code here and/or call default

	TRACE2("CAddressView::OnLButtonDblClk()   offset(%d, %d)", offsetX, offsetY);
	TRACE2("   point(%d, %d)\n", point.x, point.y);

	i = 0;
	// find table entry with given point coords
	for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
	{
		if (ptrTable->m_selected
			&& point.y <= ptrTable->m_drawCoord[1]
			&& point.y >= ptrTable->m_drawCoord[3]
			&& point.x >= ptrTable->m_drawCoord[0]
			&& point.x <= ptrTable->m_drawCoord[2])
		{
			TRACE1("   selected: %s\n", ptrTable->m_name);
			pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
			((CListBox *)pCBox)->SetCaretIndex(i, FALSE);
			pMainFrame->OnNameEditEntry();
			return;
	
		}
		i++;
	}


	CScrollView::OnLButtonDblClk(nFlags, point);
}


void CAddressView::OnLButtonUp(UINT nFlags, CPoint point) 
// sets position of first label
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	//CScrollView::OnLButtonUp(nFlags, point);


	// TODO: Add your message handler code here and/or call default
	CPoint sp = GetScrollPosition();
	int i;
	int offsetX;
	int offsetY;
	CAddressTable *ptrTable;
	CWnd *pCBox;
	MSG msg;

	// if double click don't do label update
	Sleep(::GetDoubleClickTime());
	if (::PeekMessage(&msg, NULL, WM_LBUTTONDBLCLK, WM_LBUTTONDBLCLK, PM_NOREMOVE))
	{
		return;
	}

	CDC *pDC;
	pDC = GetDC( );

	// CScrollView changes the viewport origin and mapping mode.
	// It's necessary to convert the point from device coordinates
	// to logical coordinates, such as are stored in the document.
	CClientDC dc(this);
	OnPrepareDC(&dc);
	dc.DPtoLP(&point);

	if (nFlags == MK_CONTROL || nFlags == MK_SHIFT)
	{
		i = 0;
		// find table entry with given point coords
		for (ptrTable = pDoc->m_ptrTableList; ptrTable != NULL; ptrTable = ptrTable->m_next)
		{
			if (ptrTable->m_selected
				&& point.y <= ptrTable->m_drawCoord[1]
				&& point.y >= ptrTable->m_drawCoord[3]
				&& point.x >= ptrTable->m_drawCoord[0]
				&& point.x <= ptrTable->m_drawCoord[2])
			{
				pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).GetDlgItem(IDC_NAME);
				((CListBox *)pCBox)->SetCaretIndex(i, FALSE);
				((CListBox *)pCBox)->SetSel(i, FALSE);
				pMainFrame->OnName();
				return;
		
			}
			i++;
		}

		return;
	}

	// convert from screen to logical coords
	//point.x = point.x * 100 / pDC->GetDeviceCaps(LOGPIXELSX); 
	//point.y = point.y * 100 / pDC->GetDeviceCaps(LOGPIXELSY);
	//sp.x = sp.x * 100 / pDC->GetDeviceCaps(LOGPIXELSX);
	//sp.y = sp.y * 100 / pDC->GetDeviceCaps(LOGPIXELSY);

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (!((CButton *)pCBox)->GetCheck())
	{
		return;
	}


	offsetX = pDoc->m_physicalOffsetX + point.x;// + sp.x;
	//offsetY = pDoc->m_physicalOffsetY + point.y - 50 - sp.y;
	offsetY = pDoc->m_physicalOffsetY - point.y - pDoc->m_labelTopMargin;

	if (offsetX / pDoc->m_labelColumnWidth < pDoc->m_labelDimension[0]
		&& offsetY / pDoc->m_labelHeight < pDoc->m_labelDimension[1])
	// within bounds of first page
	{
		pDoc->m_firstLabel = ((offsetX / pDoc->m_labelColumnWidth) * pDoc->m_labelDimension[1]) + (offsetY / pDoc->m_labelHeight);
	}

	TRACE2("CAddressView::OnLButtonUp(): %d %d   ", point.x, point.y);
	TRACE2("Scroll(%d, %d)\n", sp.x, sp.y);
	TRACE3("first label = %d    offsetXY(%d, %d)\n", pDoc->m_firstLabel, offsetX, offsetY);
	
	pDoc->rpComputeDrawPositions();

	CScrollView::OnLButtonUp(nFlags, point);
}

void CAddressView::OnRButtonUp(UINT nFlags, CPoint point)
// sets position of last return address on label
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	// TODO: Add your message handler code here and/or call default
	CPoint sp = GetScrollPosition();
	int offsetX;
	int offsetY;
	int page;
	CWnd *pCBox;

	CDC *pDC;
	pDC = GetDC( );

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_LABEL);
	if (!((CButton *)pCBox)->GetCheck())
	{
		return;
	}

	pCBox = (CComboBox*)(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).GetDlgItem(IDC_RETURN);
	if (!((CButton *)pCBox)->GetCheck())
	{
		return;
	}

	// CScrollView changes the viewport origin and mapping mode.
	// It's necessary to convert the point from device coordinates
	// to logical coordinates, such as are stored in the document.
	CClientDC dc(this);
	OnPrepareDC(&dc);
	dc.DPtoLP(&point);

	offsetX = pDoc->m_physicalOffsetX + point.x;
	page = -point.y / pDoc->m_printableHeight;
	offsetY = -page * (pDoc->m_printableHeight - pDoc->m_labelHeight * pDoc->m_labelDimension[1])
		+ pDoc->m_physicalOffsetY - point.y - pDoc->m_labelTopMargin;

	pDoc->m_lastLabel = ((offsetX / pDoc->m_labelColumnWidth) * pDoc->m_labelDimension[1])
		+ (offsetY / pDoc->m_labelHeight) + (page * 2 * pDoc->m_labelDimension[1]);

	//offsetX = pDoc->m_physicalOffsetX + point.x + sp.x;
	//page = (point.y - sp.y) / pDoc->m_printableHeight;
	//offsetY = -page * (pDoc->m_printableHeight - 1000)
	//	+ pDoc->m_physicalOffsetY - 50 + point.y - sp.y;

	//if (offsetX / 278 < 3)
	//{
	//	pDoc->m_lastLabel = ((offsetX / 278) * 10) + (offsetY / 100) + (page * 20);
	//}

	TRACE2("CAddressView::OnRButtonUp(): point(%d, %d)   ", point.x, point.y);
	//TRACE2("scroll(%d, %d)\n", sp.x, sp.y);
	TRACE3("last label = %d    offsetXY(%d, %d)\n", pDoc->m_lastLabel, offsetX, offsetY);

	pDoc->rpComputeDrawPositions();

	CScrollView::OnRButtonUp(nFlags, point);
}

void CAddressView::OnFilePrintPreview() 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).ShowWindow(SW_HIDE);
	(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).ShowWindow(SW_HIDE);
	
	CView::OnFilePrintPreview();
	return;
}

void CAddressView::OnEndPrintPreview(CDC* pDC, CPrintInfo* pInfo, POINT point, CPreviewView* pView) 
{
	CMainFrame *pMainFrame = (CMainFrame *)GetTopLevelFrame();
	if (!pMainFrame || !pMainFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
	{
		return;
	}

	CAddressDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	(((CMainFrame *)pMainFrame)->m_wndDlgBarTop).ShowWindow(SW_SHOW);
	(((CMainFrame *)pMainFrame)->m_wndDlgBarLeft).ShowWindow(SW_SHOW);
	
	CScrollView::OnEndPrintPreview(pDC, pInfo, point, pView);
}
