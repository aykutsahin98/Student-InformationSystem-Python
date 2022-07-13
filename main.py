
import pandas as pd
import openpyxl

course_list = ['Matematik', 'Fizik', 'Linear Cebir', 'Türkçe']
student_list = []


def choose_course():
    index = 1
    for course in course_list:
        print(f"{index}- {course}")
        index += 1

    choice = input('Lütfen aşağıdaki derslerden birisini numarala ile seçiniz: ')

    if choice not in ['1', '2', '3', '4']:
        print("Lütfen geçerli bir değer girin!")
    else:
        chosen_course = course_list[int(choice) - 1]
        print(f"{chosen_course} dersi seçildi!")
        return course_list[int(choice) - 1]


def enroll_user():
    first_name = input("İsim: ")
    last_name = input("Soyisim: ")
    student_id = input("Öğrenci No: ")
    grade = int(input("Not (0-100): "))
    letter_grade = calculate_letter_grade(grade)

    isPassed = False
    if letter_grade != 'F':
        isPassed = True

    student = {
        'ID': student_id,
        'İsim': first_name,
        'Soyisim': last_name,
        'Harf Notu': letter_grade,
        'Geçti': isPassed
    }
    student_list.append(student)

    print("Yeni öğrenci sisteme eklendi!\n")


def export_into_file(course_name):
    df = pd.DataFrame(student_list)
    df = df.style.set_properties(**{'text-align': 'center'})
    df.to_excel(f'{course_name}-Öğrenci-Listesi.xlsx', index=False, header=True)


def calculate_letter_grade(grade):
    if grade > 85:
        return 'A'
    elif grade > 70:
        return 'B'
    elif grade > 55:
        return 'C'
    elif grade > 50:
        return 'D'

    return 'F'


def exit_program(course_name):
    print("Dosya işleniyor...")
    export_into_file(course_name)
    print("Tamamlandı!")
    exit()


def main():
    course_name = choose_course()

    while True:
        command = input("Öğrenci kaydı yapmak istiyorsanız Y, çıkış yapmak istiyorsanız Q: ")

        if command == 'Q' or command == 'q':
            if len(student_list) > 0:
                exit_program(course_name)
            break
        elif command == 'Y' or command == 'y':
            enroll_user()
        else:
            print("Geçersiz komut!")


if __name__ == '__main__':
    main()
