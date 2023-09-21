import openpyxl as xl
import random


# Define Member class
class Member:
    def __init__(self, name, lion, drum, box, bench):
        self.name = name
        self.lion = lion
        self.drum = drum
        self.box = box
        self.bench = bench

    def __str__(self):
        return f"{self.name}, {self.lion}, {self.drum}, {self.box}, {self.bench}"

    def name(self):
        return self.name

    def lion(self):
        return self.lion

    def drum(self):
        return self.lion

    def box(self):
        return self.box

    def bench(self):
        return self.bench


def text_to_boolean(text):
    if "Y" in text or "y" in text or "M" in text or "m" in text:
        return True
    else:
        return False


def generate_members(availability_spreadsheet):
    # Load Workbooks
    availability_book = xl.load_workbook(filename=availability_spreadsheet)

    # Get Proper Pages
    availability_page = availability_book["Availability"]

    # Set of members
    member_set = set()

    # Create Members and add to member_set
    for r in range(2, availability_page.max_row + 1):
        member_info = availability_page[r]
        member_set.add(Member(member_info[0].value, text_to_boolean(member_info[1].value),
                              text_to_boolean(member_info[2].value), text_to_boolean(member_info[3].value),
                              text_to_boolean(member_info[4].value)))

    return member_set

def generate_lion_members(availability_spreadsheet):
    # Load Workbooks
    availability_book = xl.load_workbook(filename=availability_spreadsheet)

    # Get Proper Pages
    availability_page = availability_book["Availability"]

    # Set of members
    lion_set = set()

    # Create Members and add to member_set
    for r in range(2, availability_page.max_row + 1):
        member_info = availability_page[r]
        if text_to_boolean(member_info[1].value):
            lion_set.add(member_info[0].value)

    return lion_set


def generate_drum_members(availability_spreadsheet):
    # Load Workbooks
    availability_book = xl.load_workbook(filename=availability_spreadsheet)

    # Get Proper Pages
    availability_page = availability_book["Availability"]

    # Set of members
    drum_set = set()

    # Create Members and add to member_set
    for r in range(2, availability_page.max_row + 1):
        member_info = availability_page[r]
        if text_to_boolean(member_info[2].value):
            drum_set.add(member_info[0].value)

    return drum_set


def generate_box_members(availability_spreadsheet):
    # Load Workbooks
    availability_book = xl.load_workbook(filename=availability_spreadsheet)

    # Get Proper Pages
    availability_page = availability_book["Availability"]

    # Set of members
    box_set = set()

    # Create Members and add to member_set
    for r in range(2, availability_page.max_row + 1):
        member_info = availability_page[r]
        if text_to_boolean(member_info[3].value):
            box_set.add(member_info[0].value)

    return box_set


def generate_bench_members(availability_spreadsheet):
    # Load Workbooks
    availability_book = xl.load_workbook(filename=availability_spreadsheet)

    # Get Proper Pages
    availability_page = availability_book["Availability"]

    # Set of members
    bench_set = set()

    # Create Members and add to member_set
    for r in range(2, availability_page.max_row + 1):
        member_info = availability_page[r]
        if text_to_boolean(member_info[4].value):
            bench_set.add(member_info[0].value)

    return bench_set


def check_used(row, name, rotation_page):
    for c in range(2, 14):
        if rotation_page.cell(row=row, column=c).value == name:
            return True
    return False


# Load Workbooks
equipment_rotation_book = xl.load_workbook(filename="Equipment Rotation.xlsx")
availability_book = xl.load_workbook(filename="Equipment Rotation Availability.xlsx")
equipment_rotation_page = equipment_rotation_book["Sept 23"]
availability_page = availability_book["Availability"]

for row in range(2, 12):
    lion_used = False
    senior_used = False
    bench_set = generate_bench_members("Equipment Rotation Availability.xlsx")
    drum_set = generate_drum_members("Equipment Rotation Availability.xlsx")
    box_set = generate_box_members("Equipment Rotation Availability.xlsx")
    lion_set = generate_lion_members("Equipment Rotation Availability.xlsx")

    for col in range(13, 1, -1):
        member = "^"
        if equipment_rotation_page.cell(row=row, column=col).value is not None:
            pass
        elif equipment_rotation_page.cell(row=1, column=col).value == "Benches":
            member = random.choice(list(bench_set))
            while check_used(row, member, equipment_rotation_page):
                member = random.choice(list(bench_set))
        elif "Drum" in equipment_rotation_page.cell(row=1, column=col).value:
            member = random.choice(list(drum_set))
            while check_used(row, member, equipment_rotation_page):
                member = random.choice(list(drum_set))
        elif "Box" in equipment_rotation_page.cell(row=1, column=col).value:
            member = random.choice(list(box_set))
            while check_used(row, member, equipment_rotation_page):
                member = random.choice(list(box_set))
        else:
            member = random.choice(list(lion_set))
            while check_used(row, member, equipment_rotation_page):
                member = random.choice(list(lion_set))

        if "Lions" in member:
            for i in range(2, 23):
                if "Lions" in availability_page.cell(row=i, column=1).value:
                    bench_set.discard(availability_page.cell(row=i, column=1).value)
                    drum_set.discard(availability_page.cell(row=i, column=1).value)
                    box_set.discard(availability_page.cell(row=i, column=1).value)
                    lion_set.discard(availability_page.cell(row=i, column=1).value)

        if "Senior" in member:
            for i in range(2, 23):
                if "Senior" in availability_page.cell(row=i, column=1).value:
                    bench_set.discard(availability_page.cell(row=i, column=1).value)
                    drum_set.discard(availability_page.cell(row=i, column=1).value)
                    box_set.discard(availability_page.cell(row=i, column=1).value)
                    lion_set.discard(availability_page.cell(row=i, column=1).value)
        if member != "^":
            equipment_rotation_page.cell(row=row, column=col, value=member)
        equipment_rotation_book.save('Equipment Rotation.xlsx')
