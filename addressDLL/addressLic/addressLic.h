/*
	this header is not required for this dll
	to compile; but is required for the app
	that will use this dll
*/

#define ADDRESS_ERROR_LICENSE_PIPE_FILE -100
#define ADDRESS_ERROR_NO_ETHERNET_CARD -101
#define ADDRESS_ERROR_INVALID_LICENSE -102

#ifndef CPPDLL_H
#define CPPDLL_H

//prevent the function name from being mangled
extern "C" 
void rpValidLicense(char *code);

#endif //CPPDLL_H