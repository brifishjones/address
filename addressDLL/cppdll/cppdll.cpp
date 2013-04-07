#include "cppdll.h"

//prevent function name from being mangled
extern "C" 
int cube(int num) {
	return num * num * num;
}


extern "C" 
int rpValidLicense(void)
{
	return 99;
}