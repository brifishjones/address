#include <iostream>

#include "cppdll.h"

using namespace std;

int main() {
	int num;

	cout << "enter any number: ";
	cin >> num;

	cout << "cube of number is " 
		 << cube(num) << endl;

	return 0;
}