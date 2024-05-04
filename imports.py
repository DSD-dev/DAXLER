from telethon.sync import TelegramClient
from telethon.tl.functions.channels import GetParticipantsRequest
from telethon.tl.types import ChannelParticipantsSearch
from telethon.errors.rpcerrorlist import FloodWaitError, ChatAdminRequiredError
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from termcolor import colored