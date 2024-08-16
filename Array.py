import pandas as pd

class Student:
    def __init__(self, name, id_number, branch, age, year):
        self.name = name
        self.id_number = id_number
        self.branch = branch
        self.age = age
        self.year = year

    def __repr__(self):
        return f"{self.name} (ID: {self.id_number}, Branch: {self.branch}, Age: {self.age}, Year: {self.year})"

def insert_at_beginning(array, student):
    array.insert(0, student)

def insert_at_end(array, student):
    array.append(student)

def insert_at_index(array, index, student):
    if 0 <= index <= len(array):
        array.insert(index, student)
    else:
        print("Index out of range.")

def delete_at_index(array, index):
    if 0 <= index < len(array):
        array.pop(index)
    else:
        print("Index out of range.")

def update_at_index(array, index, student):
    if 0 <= index < len(array):
        array[index] = student
    else:
        print("Index out of range.")

def search_element(array, id_number):
    for i, student in enumerate(array):
        if student.id_number == id_number:
            return i
    return -1

def sort_array(array):
    array.sort(key=lambda s: s.id_number)

def display_array(array):
    for student in array:
        print(student)

def export_to_excel(array):
    # Convert list of students to a DataFrame
    df = pd.DataFrame([{
        'Name': student.name,
        'ID': student.id_number,
        'Branch': student.branch,
        'Age': student.age,
        'Year': student.year
    } for student in array])
    
    # Save DataFrame to Excel
    df.to_excel('students.xlsx', index=False)
    print("Data exported to 'students.xlsx'.")

def main():
    array = []
    n = int(input("Enter the number of students: "))

    for _ in range(n):
        name = input("Enter student name: ")
        id_number = int(input("Enter student ID number: "))
        branch = input("Enter student branch: ")
        age = int(input("Enter student age: "))
        year = int(input("Enter student year of studying: "))
        array.append(Student(name, id_number, branch, age, year))

    while True:
        choice = int(input(
            "\nChoose an operation:\n"
            "1. Insert at beginning\n"
            "2. Insert at end\n"
            "3. Insert at index\n"
            "4. Delete at index\n"
            "5. Update at index\n"
            "6. Search for student by ID\n"
            "7. Sort array by ID\n"
            "8. Display students\n"
            "9. Export to Excel\n"
            "10. Exit\n"
            "Enter your choice: "
        ))

        if choice == 1:
            student = get_student_input()
            insert_at_beginning(array, student)
        elif choice == 2:
            student = get_student_input()
            insert_at_end(array, student)
        elif choice == 3:
            index = int(input("Index to insert at: "))
            student = get_student_input()
            insert_at_index(array, index, student)
        elif choice == 4:
            index = int(input("Index to delete: "))
            delete_at_index(array, index)
        elif choice == 5:
            index = int(input("Index to update: "))
            student = get_student_input()
            update_at_index(array, index, student)
        elif choice == 6:
            id_number = int(input("Student ID to search for: "))
            index = search_element(array, id_number)
            if index != -1:
                print(f"Student found: {array[index]}")
            else:
                print("Student not found.")
        elif choice == 7:
            sort_array(array)
            print("Array sorted by student ID.")
        elif choice == 8:
            display_array(array)
        elif choice == 9:
            export_to_excel(array)
        elif choice == 10:
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")

def get_student_input():
    name = input("Enter student name: ")
    id_number = int(input("Enter student ID number: "))
    branch = input("Enter student branch: ")
    age = int(input("Enter student age: "))
    year = int(input("Enter student year of studying: "))
    return Student(name, id_number, branch, age, year)

if __name__ == "__main__":
    main()
