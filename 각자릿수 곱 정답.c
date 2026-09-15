int main()

{
    int a,b;
    int temp;
    scanf("%d", &a);
    scanf("%d", &b);

    temp = b;
    while(temp > 0) //(temp !=0)
    {
        printf("%d\n", a*(temp%10));
        temp/=10;
    }
    printf("%d", a*b);
}

