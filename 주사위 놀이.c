#include <stdio.h>

int main() {
    srand(time(NULL)); 

    for (int i = 0; i < 10; i++) {
        int dice = rand() % 6 + 1; 
        printf("주사위 %d번째 결과: %d\n", i + 1, dice);
    }

    return 0;
}
