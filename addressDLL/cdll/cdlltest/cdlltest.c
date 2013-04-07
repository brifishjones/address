#include <stdio.h>

#include "cdll.h"

int main() {
	int num;
	
	printf("enter a number: ");
	scanf("%d", &num);
	printf("square is %d\n", square(num));

	return 0;
}