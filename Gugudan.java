import java.util.Scanner;

public class Gugudan {
    private static void printGugudan(int dan) {
        for (int i = 1; i <= 9; i++) {
            System.out.printf("%d x %d = %d%n", dan, i, dan * i);
        }
    }

    private static void printAllGugudan() {
        for (int dan = 2; dan <= 9; dan++) {
            System.out.printf("%n[%d단]%n", dan);
            printGugudan(dan);
        }
    }

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.print("출력할 단을 입력하세요(2~9, 전체는 Enter): ");
        String input = scanner.nextLine().trim();

        if (input.isEmpty()) {
            printAllGugudan();
            return;
        }

        try {
            int dan = Integer.parseInt(input);
            if (dan < 2 || dan > 9) {
                System.out.println("2부터 9 사이의 숫자를 입력하세요.");
                return;
            }
            printGugudan(dan);
        } catch (NumberFormatException e) {
            System.out.println("숫자를 입력하세요.");
        }
    }
}
