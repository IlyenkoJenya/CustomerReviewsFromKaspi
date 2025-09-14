from dotenv import load_dotenv
import os

load_dotenv()

BOT_TOKEN = os.getenv("bot")
CHAT_ID = os.getenv("chatId")
CHAT_ID_SERVICE = os.getenv("chatId_service")
TOKEN_FIRST = os.getenv("tokenFIRST")
TOKEN_SECOND = os.getenv("tikenSECOND")
ID_FIRST = os.getenv("idFIRST")
ID_SECOND = os.getenv("idSECOND")