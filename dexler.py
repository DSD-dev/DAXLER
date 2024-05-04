from imports import *
import os
import pyfiglet
import random
import time

os.system("clear")

def write_to_file(file_name, message):
    with open(file_name, 'w') as f:
        f.write(message)

def read_from_file(file_name):
    try:
        with open(file_name, 'r') as f:
            return f.read().strip()
    except FileNotFoundError:
        print(f"File {file_name} not found.")
        exit()

def random_color():
    colors = ['red', 'green', 'yellow', 'blue', 'magenta', 'cyan', 'white']
    return random.choice(colors)

dexler_text = pyfiglet.figlet_format("Dexler", font="epic")
print(colored(dexler_text, "red"))

version_text = "version 1.0"
for char in version_text:
    color = random_color()
    print(colored(char, color), end="", flush=True)
    time.sleep(0.1)
print()
created_by_text = "Created by νέος"
for char in created_by_text:
    color = random_color()
    print(colored(char, color), end="", flush=True)
    time.sleep(0.1)
print()
phone_number = input("Phone number: ")

api_id_file = 'api_id.txt'
api_hash_file = 'api_hash.txt'
try:
    api_id = read_from_file(api_id_file)
    api_hash = read_from_file(api_hash_file)
except FileNotFoundError:
    api_id = input("API ID: ")
    api_hash = input("API Hash: ")
    write_to_file(api_id_file, api_id)
    write_to_file(api_hash_file, api_hash)

client = TelegramClient('session_name', api_id, api_hash)
client.connect()

if not client.is_user_authorized():
    client.send_code_request(phone_number)
    client.sign_in(phone_number, input('Enter the code from SMS: '))

group_name = input("Enter the group name: ")
group_username = input("Enter the group link: ")

if "https://t.me/" in group_username:
    group_username = group_username.replace("https://t.me/", "@")

try:
    group_entity = client.get_entity(group_username if group_username else group_name)
except ValueError:
    print("Group not found.")
    exit()

limit = 250

print(f"Parsing group {group_entity.title}; participants count: {group_entity.participants_count}...")

all_participants = []
try:
    offset = 0
    while True:
        participants = client(GetParticipantsRequest(group_entity, ChannelParticipantsSearch(''), offset, limit, 0))
        if not participants.users:
            break
        all_participants.extend(participants.users)
        offset += len(participants.users)
except FloodWaitError as e:
    print(f"FloodWait error: Please wait {e.seconds} seconds and try again.")
except ChatAdminRequiredError:
    print("Admin rights required to get the list of participants.")
    exit()

if not os.path.exists('db'):
    os.makedirs('db')

os.chdir('db')

wb = Workbook()
ws = wb.active
ws.append(["ID", "", "NAME", "", "NUMBER", "", "USERNAME", "", "LINK"])
id_style = NamedStyle(name="ID_Style")
id_style.number_format = '0' * 10
wb.add_named_style(id_style)

excluded_ids = {5118537664, 1556857003, 1528181771}

for participant in all_participants:
    if participant.id not in excluded_ids:
        full_name = participant.first_name + " " + participant.last_name if participant.last_name else participant.first_name
        phone = participant.phone if participant.phone else ""
        username = participant.username if participant.username else ""

        group_link = f"https://t.me/{group_entity.username}" if group_entity.username and len(ws['I']) == 1 else ""
        ws.append([participant.id, "", full_name, "", phone, "", username, "", group_link])

file_name = group_entity.username if group_entity.username else group_name
file_name += "_members.xlsx"
wb.save(file_name)

print(colored(f"Parsing group {group_entity.title} finished. Database name {file_name}; number of rows {len(ws['A'])}", "green"))

choice = input("Want to parse another group? (y/n): ")
if choice.lower() == 'y':
    os.system("python3 dexler.py")
else:
    print("Goodbye!")