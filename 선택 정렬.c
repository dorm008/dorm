#include <stdio.h>

void selection_sort(int arr[], int n) {
    for (int i = 0; i < n - 1; i++) {
        int min_idx = i;
        // 남은 부분에서 최솟값 찾기
        for (int j = i + 1; j < n; j++) {
            if (arr[j] < arr[min_idx]) {
                min_idx = j;
            }
        }
        // 최솟값을 맨 앞과 교환
        int temp = arr[i];
        arr[i] = arr[min_idx];
        arr[min_idx] = temp;
    }
}

int main() {
    int data[] = {5, 3, 8, 4, 2};
    int n = sizeof(data) / sizeof(data[0]);

    selection_sort(data, n);

    printf("선택 정렬 결과: ");
    for (int i = 0; i < n; i++) {
        printf("%d ", data[i]);
    }
    printf("\n");

    return 0;
}
