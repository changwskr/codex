def print_gugudan(dan):
    for i in range(1, 10):
        print(f"{dan} x {i} = {dan * i}")


def print_all_gugudan():
    for dan in range(2, 10):
        print(f"\n[{dan}단]")
        print_gugudan(dan)


def main():
    value = input("출력할 단을 입력하세요(2~9, 전체는 Enter): ").strip()

    if not value:
        print_all_gugudan()
        return

    if not value.isdigit():
        print("숫자를 입력하세요.")
        return

    dan = int(value)
    if dan < 2 or dan > 9:
        print("2부터 9 사이의 숫자를 입력하세요.")
        return

    print_gugudan(dan)


if __name__ == "__main__":
    main()
