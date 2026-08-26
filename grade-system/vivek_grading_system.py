"""
grade_system.py
Asks the user for a mark (0-100) and prints the matching letter grade.

Grading Scale:
    90 - 100 -> A
    80 - 89  -> B
    70 - 79  -> C
    60 - 69  -> D
    below 60 -> E
"""


def get_grade(mark):
    """Return the letter grade for a valid mark between 0 and 100."""
    if mark >= 90:
        return "A"
    elif mark >= 80:
        return "B"
    elif mark >= 70:
        return "C"
    elif mark >= 60:
        return "D"
    else:
        return "E"


def get_valid_mark():
    """Keep asking the user until a valid numeric mark between 0 and 100 is entered."""
    while True:
        user_input = input("Enter a mark (0-100): ").strip()

        try:
            mark = float(user_input)
        except ValueError:
            print(f"'{user_input}' is not a valid number. Please try again.\n")
            continue

        if mark < 0 or mark > 100:
            print(f"{mark} is out of range. Please enter a value between 0 and 100.\n")
            continue

        return mark


def main():
    mark = get_valid_mark()
    grade = get_grade(mark)

    # Show whole numbers cleanly (e.g. 90 instead of 90.0)
    display_mark = int(mark) if mark == int(mark) else mark

    print(f"\nMark entered: {display_mark}")
    print(f"Grade: {grade}")


if __name__ == "__main__":
    main()