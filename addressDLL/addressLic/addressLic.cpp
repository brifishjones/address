// addressLic.cpp : Defines the entry point for the DLL application.
//
#include "stdafx.h"
#include "addressLic.h"
//#include <stdio.h>

// progID is the usually the CLSID of the program without braces and dashes
char progID[48] = "F3CAEE0F4D244800A7469B7080362802";

#define ADDRESS_LICENSE "ibcfjamhglqskdne"

// Russ Leggett			00-04-5A-66-F0-99	bqux dtwc zjrf pkea
// Joanna Sinnwell		00-00-86-46-D1-79	jyin bsgk tcmo awde
// Bill Fisher			00-50-BF-75-01-CF	ibcf jamh glqs kdne
 
//prevent function name from being mangled
extern "C" 
void rpValidLicense(char *license)
{
	//char code[128];
	int i;
	int j;
	int k;
	double a;
	double a0;
	double b;
	double c;
	double r;
	char s[16];
	static int vCount = 0;

	if (license[0] != 'A')
	// return license
	{
		sprintf(license, ADDRESS_LICENSE);
		return;
	}

	//sprintf(code, "%s", license);
	
	// perform license validation using the same encoding routine found in
	// address.exe [BOOL CAddressApp::rpLicense()]

	//k = (int)code[93] - 48;
	//r = 3.9 + (double)((int)code[k * 7 + 3] - 48) * 0.01
	//	+ (double)((int)code[k * 7 + 5] - 48) * 0.001
	//	+ (double)((int)code[k * 7 + 8] - 48) * 0.0001
	//	+ (double)((int)code[k * 7 + 13] - 48) * 0.00001
	//	+ (double)((int)code[k * 7 + 15] - 48) * 0.000001;
	//a = (double)((int)code[k * 8 + 3] - 48) * 0.1
	//	+ (double)((int)code[k * 8 + 5] - 48) * 0.01
	//	+ (double)((int)code[k * 8 + 8] - 48) * 0.001
	//	+ (double)((int)code[k * 8 + 13] - 48) * 0.0001;

	// return verification code
	// will be called 8 times by address.exe
	// part of the program ID (usually the CLSID) is used for r and a0
	if (vCount >= 8)
	// vCount should never reach 8
	{
		for (i = 0; i < 126; i++)
		{
			license[i] = '0';
		}
		license[126] = '\0';
		return;
	}

	// r and a0 is a blend of progID and the current date code imbedded in license
	k = (int)license[93] - 48;
	r = 3.9 + (double)((int)progID[vCount] % 10) * 0.01
		+ (double)((int)license[k * 7 + 5] - 48) * 0.001
		+ (double)((int)progID[16 + vCount] % 10) * 0.0001
		+ (double)((int)license[k * 7 + 13] - 48) * 0.00001
		+ (double)((int)progID[24 + vCount] % 10) * 0.000001;
	a0 = (double)((int)progID[8 + vCount] % 10) * 0.1
		+ (double)((int)license[k * 7 + 8] - 48) * 0.01
		+ (double)((int)progID[vCount] % 10) * 0.001
		+ (double)((int)license[k * 7 + 3] - 48) * 0.0001
		+ (double)((int)progID[16 + vCount] % 10) * 0.00001
		+ (double)((int)license[k * 7 + 15] - 48) * 0.000001;

	//r = 3.9 + (double)((int)license[k * 7 + 3] - 48) * 0.01
	//	+ (double)((int)license[k * 7 + 5] - 48) * 0.001
	//	+ (double)((int)license[k * 7 + 8] - 48) * 0.0001
	//	+ (double)((int)license[k * 7 + 13] - 48) * 0.00001
	//	+ (double)((int)license[k * 7 + 15] - 48) * 0.000001;
	//a0 = (double)((int)license[k * 8 + 3] - 48) * 0.1
	//	+ (double)((int)license[k * 8 + 5] - 48) * 0.01
	//	+ (double)((int)license[k * 8 + 8] - 48) * 0.001
	//	+ (double)((int)license[k * 8 + 13] - 48) * 0.0001;

	if (a0 < 0.0001)
	{
		a0 = 0.0001;
	}

	for (i = 0; i < 21; i++)
	{
		a = a0; // always reset a to prevent rounding error
		for (j = 0; j < 21 - ((i * 4) % 15); j++)
		{
			b = r * a;
			c = 1.0 - a;
			a = b * c;
		}
		//k = (int)(a * 1000000.0);
		sprintf(s, "%lf", a);
		license[i * 6] = s[2];
		license[i * 6 + 1] = s[3];
		license[i * 6 + 2] = s[4];
		license[i * 6 + 3] = s[5];
		license[i * 6 + 4] = s[6];
		license[i * 6 + 5] = s[7];
		//sprintf(&license[i * 6], "%06d", k);
	}
	license[126] = '\0';

	vCount++;

	//k++;

	return;

		
/***
	//int i;
	//int j;
	//char sLicense[24];

	// initialize license
	sprintf(m_license, "0123456789ABCDEF");
	//sprintf(*license, "0123456789ABCDEF");

	// check for valid license
	if (rpLicensePipe() == FALSE && rpLicenseFile() == FALSE)
	{
		//AfxMessageBox("Unable to run address application!     \nUnable to initiate licensing     \nroutines on this computer.     ");
		return ADDRESS_ERROR_LICENSE_PIPE_FILE;
	}

	if (rpEncode() == FALSE)
	{
		if (strcmp(m_license, "0123456789ABCDEF") == 0)
		{
			//AfxMessageBox("Unable to run address application!     \nAn ethernet card must be installed on     \nthis computer for licensing purposes.     ");
			return ADDRESS_ERROR_NO_ETHERNET_CARD;
		}
		else
		{
			//j = 0;
			//for (i = 0; i < 16; i++)
			//{
			//	sLicense[j] = m_license[i];
			//	if (i % 4 == 3)
			//	{
			//		j++;
			//		sLicense[j] = ' ';
			//	}
			//	j++;
			//}
			//sLicense[j] = '\0';
			//sprintf(cmd, "Unable to run address application due to licensing error.\n%s\nValid license is: %s", m_code[0], sLicense);
			//AfxMessageBox(cmd);
			//licenseDlg.m_license = (CString)sLicense;
			//licenseDlg.DoModal();
			sprintf(*license, "%s", m_license);
			return ADDRESS_ERROR_INVALID_LICENSE;
		}
	}
	//sprintf(*license, "%s", m_license);
	return TRUE;
	***/

}

