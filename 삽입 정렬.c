#include <stdio.h>

void insertion_sort(int arr[], int n) {
    for (int i = 1; i < n; i++) {
        int key = arr[i];
        int j = i - 1;
        while (j >= 0 && arr [j] > key) {
            arr [j + 1] = arr[j];
            j--;
        }
        arr[j + 1] = key;
    }
}

int main() {
    int data[] = {5, 3, 8, 4, 2};
    int n = sizeof(data) / sizeof(data[0]);

    insertiom_sort(data, n);

    printf("삽입 정렬 결과: ");
    for (int i = 0; i < n; i++) {
        printf("%d ", data[i]);
    }
    printf("\n");

    return 0;
}