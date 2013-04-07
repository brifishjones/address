/*
	this header is not required for this dll
	to compile; but is required for the app
	that will use this dll
*/

#ifndef CPPDLL_H
#define CPPDLL_H

//computes the square of the number
//prevent the function name from being mangled
extern "C" 
int cube(int num);


extern "C" 
int rpValidLicense(void);

#endif //CPPDLL_H