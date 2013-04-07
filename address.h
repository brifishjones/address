// address.h : main header file for the ADDRESS application
//

#if !defined(AFX_ADDRESS_H__ACB543F6_7A3E_44FE_A283_C919F40C0E0B__INCLUDED_)
#define AFX_ADDRESS_H__ACB543F6_7A3E_44FE_A283_C919F40C0E0B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CAddressApp:
// See address.cpp for the implementation of this class
//

class CAddressApp : public CWinApp
{
public:
	CAddressApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddressApp)
	public:
	virtual BOOL InitInstance();
	virtual BOOL OnIdle(LONG lCount);
	//}}AFX_VIRTUAL

// Implementation
	BOOL rpMessage(char sMessage[256]);
	//BOOL rpLicense(void);
	//BOOL rpLicenseFile(void);
	//BOOL rpLicensePipe(void);
	//BOOL rpInitializationCode(void);
	//BOOL rpInitializeRegistry(char);
	//BOOL rpValidateTemporaryLicense(void);
	//BOOL rpWriteRegistry(void);
	//BOOL rpTemporaryKey(void);
	//BOOL rpEncode(char *);
	//void rpDateCode(void);
	//BOOL rpDecode(char code[17]);
	COleTemplateServer m_server;
	BOOL m_init;
	//char m_code[10][18];
	//char m_dateCode[128];
	//char m_initializationCode[17];
	//char m_temporaryKey[17];  	// temporaryKey is valid if it is not empty string

	//CTime m_time[4];

	//int m_sessions;
	//int m_clockSync;

		// Server object for document creation
	//{{AFX_MSG(CAddressApp)
	afx_msg void OnAppAbout();
	afx_msg void OnFilePrintSetup();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDRESS_H__ACB543F6_7A3E_44FE_A283_C919F40C0E0B__INCLUDED_)
