#include <stdio.h>

int main()
{
    int outcome;

    printf("범위 안에 있는 숫자를 골라서 입력하시오0(90~100), 1(80~90), 2(70~80), 3(60~70), 4(0~59):");
    scanf("%d",&outcome)

    switch (outcome){
    case 0:
        printf("90~100\n")
        printf("A");
        break;
    }
    case 1:
        printf("80~90\n")
        printf("B");
        break;
    case 2:
        printf("70~80\n")
        printf("C");
        break;
    case 3:
        printf("60~70\n")
        printf("D");
        break;
    case 4:
        printf("0~59\n")
        printf("F");
        break;
    



}