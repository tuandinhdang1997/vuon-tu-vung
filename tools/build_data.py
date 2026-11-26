"""Sinh js/data.js từ file Excel từ vựng.

Cách chạy:
    python -I tools/build_data.py "<đường dẫn tới file .xlsx>"

Script đọc 19 sheet, sửa các lỗi đã biết trong tài liệu gốc (xem FIXES),
gắn emoji minh hoạ và ghi ra js/data.js cho web app.
"""
import io
import json
import re
import sys
from pathlib import Path

import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

OUT = Path(__file__).resolve().parent.parent / "js" / "data.js"

# (nhãn buổi, chủ điểm EN, chủ điểm VI, emoji, giọng của phiên âm)
LESSONS = [
    ("Buổi 1", "Office Issues", "Chuyện văn phòng", "🗂️", "uk"),
    ("Buổi 2", "Personnel & Hiring", "Nhân sự & Tuyển dụng", "🤝", "uk"),
    ("Buổi 3", "Business Meetings", "Họp hành", "📊", "uk"),
    ("Tổng ôn", "Finance & Budgeting", "Tài chính & Ngân sách", "💰", "uk"),
    ("Buổi 4", "Marketing & Advertising", "Tiếp thị & Quảng cáo", "📣", "uk"),
    ("Buổi 5", "Mixed Vocabulary", "Từ vựng tổng hợp", "🧩", "uk"),
    ("Buổi 6", "Sales & Purchasing", "Bán hàng & Mua sắm", "🛒", "uk"),
    ("Buổi 7", "Travel & Transportation", "Du lịch & Đi lại", "✈️", "uk"),
    ("Buổi 8", "Dining & Catering", "Ăn uống & Tiệc tùng", "🍽️", "uk"),
    ("Buổi 9", "Technology & Electronics", "Công nghệ & Điện tử", "💻", "uk"),
    ("Buổi 10", "Health & Insurance", "Sức khỏe & Bảo hiểm", "🩺", "uk"),
    ("Buổi 11", "Property & Real Estate", "Bất động sản", "🏠", "uk"),
    ("Buổi 12", "Manufacturing & Production", "Sản xuất", "🏭", "uk"),
    ("Buổi 13", "Customer Service", "Dịch vụ khách hàng", "🎧", "uk"),
    ("Buổi 14", "Health", "Sức khỏe", "💊", "uk"),
    ("Buổi 15", "Corporate Events & Entertainment", "Sự kiện công ty & Giải trí", "🎉", "us"),
    ("Buổi 16", "Shipping & Logistics", "Vận chuyển & Hậu cần", "📦", "us"),
    ("Buổi 17", "Legal & Contracts", "Pháp lý & Hợp đồng", "⚖️", "us"),
    ("Buổi 18", "Facility Maintenance & Renovations", "Bảo trì & Cải tạo", "🛠️", "us"),
]

# Sửa lỗi của tài liệu gốc: (chỉ số sheet, từ viết thường) -> các trường thay thế
FIXES = {
    (1, "manager"): {"ipa": "/ˈmænɪdʒə(r)/"},
    (1, "achievement"): {"ipa": "/əˈtʃiːvmənt/"},
    (2, "branch"): {"ipa": "/brɑːntʃ/"},
    (2, "headquarter"): {"w": "Headquarters", "ipa": "/ˌhedˈkwɔːtəz/"},
    (4, "promote"): {"ipa": "/prəˈməʊt/"},
    (5, "interpreter"): {"ipa": "/ɪnˈtɜːprətə(r)/"},
    (7, "shuttle"): {"ipa": "/ˈʃʌtl/"},
    (8, "patron"): {"vi": "Khách quen"},
    (8, "atmosphere"): {"ipa": "/ˈætməsfɪə(r)/"},
    (8, "catering service"): {
        "ipa": "/ˈkeɪtərɪŋ ˌsɜːvɪs/",
        "vi": "Dịch vụ phục vụ, cung cấp thực phẩm",
    },
    (10, "expire"): {"vi": "Hết hạn"},
    (13, "type"): {"ipa": "/taɪp/", "vi": "Đánh máy"},
    (14, "prescription"): {"ipa": "/prɪˈskrɪpʃn/"},
    (14, "security"): {"vi": "Bảo mật / An ninh"},
    (15, "rsvp"): {"ipa": "/ˌɑːr.es.viːˈpiː/"},
    (18, "system"): {"ipa": "/ˈsɪs.təm/"},
}

# Văn bản đọc riêng cho máy đọc (khi cách viết không đọc thẳng được)
SAY = {
    "rsvp": "R S V P",
    "check-up": "check up",
    "warrantee / warranty": "warranty",
    "non-disclosure agreement (nda)": "non-disclosure agreement",
}

EMOJI = {
    # văn phòng
    "colleague": "🧑‍🤝‍🧑", "equipment": "🧰", "stationery": "✏️", "cabinet": "🗄️",
    "photocopy": "📠", "stapler": "📎", "deadline": "⏰", "submit": "📤", "memo": "📝",
    "boardroom": "🏛️", "folder": "📁", "scanner": "🖨️", "ink": "🖋️", "paperwork": "📑",
    "shift": "🕘", "overtime": "🌙", "commute": "🚌", "receptionist": "💁", "workspace": "🖥️",
    "projector": "📽️", "extension": "☎️", "feedback": "💬", "document": "📄", "file": "🗂️",
    "print": "🖨️", "organize": "🗃️", "report": "📈", "briefing": "🗣️",
    "recommendation": "👍", "availability": "✅",
    # số lượng
    "one": "1️⃣", "many": "🔢", "half": "🌗", "double": "✌️", "full": "🈵", "empty": "📭",
    "first": "🥇", "last": "🏁", "whole": "⭕", "part": "🧩", "enough": "👌",
    # nhân sự
    "applicant": "🙋", "resume": "📃", "qualification": "🎓", "recruit": "🧲", "vacancy": "🪑",
    "interview": "🎤", "reference": "📇", "hire": "🤝", "manager": "🧑‍💼", "salary": "💵",
    "benefit": "🎁", "candidate": "🙋", "career": "🪜", "promotion": "📈", "resign": "🚪",
    "retire": "🏖️", "skill": "🛠️", "experience": "⭐", "position": "📍", "department": "🏬",
    "staff": "👥", "employer": "👔", "employee": "🧑‍💻", "training": "🏋️", "application": "📨",
    "requirement": "📋", "contract": "📜", "bonus": "💰", "achievement": "🏆", "background": "🎒",
    # động từ thường gặp
    "do": "✔️", "make": "🔨", "get": "🤲", "go": "🚶", "come": "👋", "take": "✊", "give": "🎁",
    "find": "🔍", "use": "🧰", "try": "💪", "know": "🧠", "think": "🤔", "feel": "💓",
    "want": "🤩", "need": "🆘", "keep": "🔒", "start": "▶️", "finish": "🏁", "learn": "📚",
    "remember": "💡",
    # họp
    "agenda": "🗒️", "participate": "🙌", "presentation": "📊", "collaborate": "🤝",
    "branch": "🏪", "headquarters": "🏢", "proposal": "💡", "negotiate": "🤝", "discuss": "💬",
    "attendee": "🧑‍🤝‍🧑", "objective": "🎯", "outcome": "🏁", "summary": "📝", "postpone": "⏳",
    "cancel": "❌", "schedule": "📅", "speaker": "🎙️", "visual aids": "🖼️", "handout": "📄",
    "conclusion": "🔚", "suggestion": "💡", "agreement": "🤝", "decision": "⚖️", "topic": "🏷️",
    "session": "🪑", "venue": "🏟️", "conduct": "🎼", "strategy": "♟️", "brainstorm": "🌩️",
    # thiên nhiên
    "sun": "☀️", "rain": "🌧️", "wind": "🌬️", "snow": "❄️", "sky": "🌤️", "tree": "🌳",
    "flower": "🌸", "water": "💧", "sea": "🌊", "mountain": "⛰️", "river": "🏞️", "animal": "🐾",
    "dog": "🐶", "cat": "🐱", "bird": "🐦", "cloud": "☁️", "warm": "🌤️", "cool": "🍃",
    "dark": "🌑", "bright": "💡",
    # tài chính
    "expense": "💸", "profit": "📈", "revenue": "💹", "audit": "🔎", "budget": "🧮",
    "invoice": "🧾", "receipt": "🧾", "transaction": "💳", "debt": "⛓️", "investment": "🌱",
    "asset": "💎", "liability": "⚠️", "bankrupt": "📉", "credit": "💳", "loan": "🏦",
    "tax": "🏛️", "interest": "➕", "fund": "🏺", "stock": "📦", "currency": "💱",
    "balance": "⚖️", "account": "👤", "cost": "🏷️", "payment": "💳", "income": "💵",
    "refund": "↩️", "financial": "💰", "statement": "📃", "forecast": "🔮", "withdraw": "🏧",
    # đời sống
    "name": "📛", "age": "🎂", "job": "💼", "friend": "🫂", "family": "👨‍👩‍👧", "address": "🏡",
    "born": "👶", "live": "🏠", "study": "📖", "hobby": "🎨", "music": "🎵", "movie": "🎬",
    "game": "🎮", "photo": "📷", "clothes": "👕", "gift": "🎁", "party": "🥳", "travel": "🧳",
    "health": "💪", "love": "❤️",
    # marketing
    "campaign": "📣", "consumer": "🛍️", "brand": "🏷️", "promote": "📢", "billboard": "🪧",
    "target": "🎯", "flyer": "📰", "competition": "🏁", "slogan": "💬", "logo": "🔰",
    "launch": "🚀", "research": "🔬", "trend": "📈", "ongoing": "🔄", "sponsor": "🤝",
    "advertise": "📺", "public relations": "🗞️", "market share": "🥧", "survey": "📋",
    "potential": "🌟", "brochure": "📖", "media": "📡", "audience": "👥", "influence": "🧲",
    "image": "🖼️", "message": "✉️", "display": "🪟", "creative": "🎨",
    # giao tiếp
    "hello": "👋", "goodbye": "👋", "yes": "✅", "no": "🚫", "please": "🙏", "thanks": "🙏",
    "sorry": "🙇", "call": "📞", "tell": "🗣️", "ask": "❓", "answer": "💬", "talk": "🗨️",
    "help": "🆘", "meet": "🤝", "send": "📤", "receive": "📥", "welcome": "🤗",
    "introduce": "🙋", "idea": "💡",
    # tổng hợp
    "perform": "🎭", "analysis": "🔬", "satisfactory": "😊", "responsibility": "✋",
    "information": "ℹ️", "director": "🧑‍💼", "timetable": "🗓️", "facilities": "🏢",
    "personnel": "👥", "advantage": "➕", "committee": "🪑", "reject": "👎", "furniture": "🛋️",
    "identification": "🆔", "anniversary": "🎂", "discounted": "🏷️", "complain": "😤",
    "complaint": "😠", "interpreter": "🗣️", "advancement": "⏫", "anticipate": "🔮",
    "arrangement": "🗓️", "relevance": "🔗", "recreational": "🎡", "expected": "⏳",
    "precaution": "🦺", "injury": "🩹", "pollution": "🏭",
    # tính từ
    "happy": "😄", "sad": "😢", "angry": "😠", "tired": "😴", "good": "👍", "bad": "👎",
    "hot": "🔥", "cold": "🥶", "big": "🐘", "small": "🐜", "new": "🆕", "old": "👴",
    "clean": "🧼", "dirty": "🦠", "easy": "😌", "hard": "🪨", "beautiful": "🌺",
    "careful": "⚠️", "quiet": "🤫", "hungry": "😋",
    # bán hàng
    "discount": "🏷️", "purchase": "🛍️", "inventory": "📦", "wholesale": "🚛", "retail": "🏪",
    "warranty": "🛡️", "order": "📝", "shipping": "🚚", "merchant": "🧑‍💼", "bargain": "🤑",
    "counter": "🧾", "shelf": "🗄️", "price tag": "🏷️", "vendor": "🧑‍🍳", "delivery": "🚚",
    "package": "📦", "coupon": "🎟️", "customer": "🧑", "seller": "🧑‍💼", "buyer": "🛒",
    "cart": "🛒", "credit card": "💳", "cash": "💵", "wallet": "👛", "supply": "🚛",
    # thời gian
    "time": "⏱️", "today": "📅", "tomorrow": "⏭️", "yesterday": "⏮️", "now": "⏺️",
    "later": "⏳", "early": "🌅", "late": "🌃", "morning": "🌄", "afternoon": "🌞",
    "evening": "🌆", "night": "🌙", "hour": "🕐", "minute": "⏲️", "second": "⏱️", "day": "🌞",
    "week": "🗓️", "month": "📆", "year": "🎆", "holiday": "🏖️",
    # du lịch
    "departure": "🛫", "arrival": "🛬", "itinerary": "🗺️", "luggage": "🧳", "booking": "📲",
    "confirmation": "✅", "delay": "⏳", "transit": "🔁", "shuttle": "🚐",
    "boarding pass": "🎫", "destination": "📍", "register": "📝", "gate": "🚪",
    "terminal": "🏢", "ticket": "🎫", "platform": "🚉", "reserve": "📌",
    "accommodation": "🏨", "vehicle": "🚗", "route": "🛣️", "customs": "🛃", "visa": "🛂",
    "passport": "📘", "passenger": "🧍", "congestion": "🚦", "trip": "🧳", "flight": "✈️",
    "cancellation": "❌", "crew": "🧑‍✈️",
    # nơi chốn
    "home": "🏡", "house": "🏠", "school": "🏫", "office": "🏢", "bank": "🏦",
    "hospital": "🏥", "hotel": "🏨", "airport": "🛫", "cinema": "🎦", "library": "📚",
    "restaurant": "🍴", "supermarket": "🛒", "gym": "🏋️", "pharmacy": "💊", "building": "🏢",
    "floor": "🧱", "garden": "🌷", "street": "🛣️", "beach": "🏖️", "city": "🏙️",
    # ăn uống
    "reservation": "📞", "menu": "📜", "beverage": "🥤", "appetizer": "🥗",
    "main course": "🍛", "utensil": "🍴", "chef": "👨‍🍳", "banquet": "🍱", "service": "🛎️",
    "dietary": "🥦", "patron": "🧑", "establishment": "🏪", "atmosphere": "🕯️",
    "dessert": "🍰", "recipe": "📖", "ingredients": "🧄", "cuisine": "🍜", "buffet": "🍢",
    "catering": "🍱", "refreshment": "🍪", "appreciated": "🙏", "tablecloth": "🧺",
    "napkin": "🧻", "complimentary": "🆓", "dine": "🍽️", "cater": "🍱",
    "catering service": "🍱", "accommodate": "🛏️", "recommend": "👍",
    "food": "🍲", "drink": "🥤", "coffee": "☕", "tea": "🍵", "eat": "😋", "breakfast": "🥐",
    "lunch": "🍱", "dinner": "🍝", "cook": "🍳", "rice": "🍚", "meat": "🥩", "fruit": "🍎",
    "bread": "🍞", "table": "🪑", "waiter": "🤵", "bill": "🧾", "delicious": "😋",
    # công nghệ
    "software": "💿", "hardware": "🔩", "install": "📥", "update": "🔄", "malfunction": "⚠️",
    "database": "🗄️", "network": "🌐", "connection": "🔌", "laptop": "💻",
    "troubleshoot": "🔧", "monitor": "🖥️", "keyboard": "⌨️", "mouse": "🖱️", "printer": "🖨️",
    "storage": "💾", "backup": "💽", "security": "🔐", "version": "🔢", "online": "🌐",
    "access": "🔑", "screen": "📺", "battery": "🔋", "charge": "🔌", "wireless": "📶",
    "device": "📱", "download": "⬇️", "upload": "⬆️",
    "store": "🏬", "shop": "🛍️", "market": "🧺", "buy": "🛒", "sell": "💰", "price": "🏷️",
    "pay": "💳", "card": "💳", "cheap": "🪙", "expensive": "💎", "sale": "🔖",
    "change": "🪙", "bag": "👜", "choice": "🔀", "total": "🧮", "size": "📏",
    # sức khỏe
    "appointment": "📅", "clinic": "🏥", "prescription": "📝", "insurance": "🛡️",
    "coverage": "☂️", "claim": "📨", "physician": "👨‍⚕️", "check-up": "🩺", "recovery": "🌈",
    "patient": "🤒", "symptom": "🤧", "medicine": "💊", "valid": "✅", "policy": "📜",
    "premium": "💎", "emergency": "🚨", "treatment": "💉", "nurse": "👩‍⚕️", "surgery": "🏥",
    "disease": "🦠", "pain": "🤕", "exercise": "🏃", "expire": "⌛", "vitamin": "🍊",
    "reimburse": "💵", "diagnosis": "🔬", "level": "📶", "contribution": "🤲",
    "conflict": "⚔️", "prevent": "🛑", "preparation": "🧺", "examine": "🔍",
    "medical": "⚕️", "operation": "🏥", "physical": "💪", "specialist": "🧑‍🔬",
    "process": "⚙️",
    # giao thông
    "car": "🚗", "bus": "🚌", "bike": "🚲", "train": "🚆", "plane": "✈️", "taxi": "🚕",
    "road": "🛣️", "corner": "↪️", "stop": "🛑", "drive": "🚘", "ride": "🏍️", "map": "🗺️",
    "park": "🅿️", "traffic": "🚦", "station": "🚉", "bridge": "🌉", "sign": "🪧", "fast": "⚡",
    # bất động sản
    "lease": "📜", "rent": "🔑", "tenant": "🧍", "landlord": "🧑‍💼", "renovate": "🔨",
    "maintenance": "🧰", "facility": "🏢", "apartment": "🏬", "utilities": "💡",
    "furnish": "🛋️", "property": "🏘️", "realtor": "🏡", "deposit": "💰",
    "residential": "🏘️", "commercial": "🏬", "location": "📍", "balcony": "🌇",
    "elevator": "🛗", "lobby": "🛎️", "repair": "🔧", "neighborhood": "🏘️", "roof": "🏠",
    "wall": "🧱", "garage": "🚗", "basement": "🪜", "hallway": "🚪",
    "key": "🔑", "door": "🚪", "window": "🪟", "box": "📦", "glass": "🥛", "cup": "☕",
    "bottle": "🍾", "money": "💵", "watch": "⌚", "glasses": "👓", "hat": "🎩",
    "shoes": "👟", "umbrella": "☂️", "mirror": "🪞", "picture": "🖼️", "light": "💡",
    "fan": "🌀",
    # sản xuất
    "factory": "🏭", "assembly line": "⚙️", "quality control": "✅", "output": "📦",
    "raw material": "🪵", "defect": "🐞", "supervisor": "🧑‍🏫", "machinery": "⚙️",
    "safety": "🦺", "mass production": "🏭", "warehouse": "🏚️", "packaging": "📦",
    "efficiency": "⚡", "standard": "📏", "labor": "👷", "tool": "🔧", "workshop": "🛠️",
    "industry": "🏭", "produce": "🏗️", "create": "✨", "build": "🧱", "design": "📐",
    "machine": "🤖", "demand": "📈", "resource": "⛏️",
    "desk": "🪑", "chair": "🪑", "computer": "🖥️", "phone": "📱", "paper": "📄",
    "pen": "🖊️", "book": "📕", "notebook": "📓", "clock": "🕰️", "lamp": "🛋️",
    "meeting": "👥", "boss": "🧑‍💼", "work": "💼", "email": "📧", "copy": "📋",
    "busy": "🐝",
    # dịch vụ khách hàng
    "satisfaction": "😊", "loyalty": "💖", "representative": "🧑‍💼", "assistance": "🤲",
    "solve": "🧩", "apology": "🙇", "help desk": "🛎️", "inquiry": "❓", "polite": "🙂",
    "resolve": "✅", "response": "↩️", "guarantee": "🛡️", "exchange": "🔄",
    "technical support": "🧑‍🔧", "review": "⭐", "helpful": "🤗", "rating": "🌟",
    "support": "🤝", "agent": "🕴️", "chat": "💬", "problem": "❗", "solution": "💡",
    "care": "💗",
    "walk": "🚶", "run": "🏃", "sit": "🪑", "stand": "🧍", "read": "📖", "write": "✍️",
    "speak": "🗣️", "listen": "👂", "sleep": "😴", "wait": "⏳", "open": "📂",
    "close": "📕", "look": "👀", "carry": "🎒", "hold": "✊", "host": "🎤", "type": "⌨️",
    # việc nhà
    "wash": "🧽", "sweep": "🧹", "dust": "🪶", "fix": "🔧", "kitchen": "🍳",
    "bedroom": "🛏️", "bathroom": "🛁", "bed": "🛏️", "soap": "🧼", "towel": "🧖",
    "brush": "🪥", "trash": "🗑️", "stairs": "🪜", "neighbor": "🏘️",
    # sự kiện
    "keynote speaker": "🎙️", "gala": "✨", "ceremony": "🎀", "reception": "🥂",
    "retirement": "🏖️", "honor": "🏅", "gathering": "🫂", "rsvp": "💌",
    "invitation": "💌", "participant": "🙋", "fundraiser": "🎗️", "seminar": "🎓",
    "conference": "🏛️", "convention": "🏟️", "registration": "📝", "enthusiastic": "🤩",
    "speech": "🎤", "networking": "🕸️", "socialize": "🥂", "unwind": "🧘",
    # vận chuyển
    "shipment": "📦", "courier": "🛵", "carrier": "🚛", "freight": "🚂", "cargo": "🚢",
    "tracking number": "🔢", "in transit": "🚚", "dispatch": "📤", "distribution": "🔀",
    "distributor": "🏬", "duty": "🧾", "tariff": "🧾", "crate": "🧰", "pallet": "🪵",
    "consignment": "📦", "deliver": "🚚", "estimate": "🧮", "express": "⚡",
    "handling fee": "💵", "overnight shipping": "🌙", "weight limit": "⚖️",
    "fragile": "🥚", "inspect": "🔍", "logistics": "🗺️",
    # pháp lý
    "clause": "📑", "terms and conditions": "📜", "comply with": "✅", "compliance": "📏",
    "renew": "🔄", "renewal": "🔄", "terminate": "✂️", "termination": "✂️", "breach": "💥",
    "obligation": "⛓️", "binding": "🔗", "non-disclosure agreement (nda)": "🤐",
    "negotiation": "🤝", "draft": "📝", "penalty": "🚩", "compensation": "💵",
    "dispute": "⚔️", "regulation": "📏", "legal counsel": "🧑‍⚖️", "signatory": "✍️",
    "invalid": "❌", "enforce": "👮", "waive": "👋", "warrantee / warranty": "🛡️",
    # cơ sở vật chất
    "renovation": "🏗️", "under construction": "🚧", "out of order": "⛔",
    "inspection": "🔍", "inspector": "🕵️", "relocate": "🚚", "relocation": "🚚",
    "restore": "♻️", "premises": "🏢", "system": "⚙️", "power outage": "🕯️",
    "plumbing": "🚰", "electrician": "👷", "technician": "🧑‍🔧", "contractor": "👷",
    "expand": "↔️", "expansion": "↔️", "upgrade": "⏫", "disruption": "🚧",
    "temporarily": "⏳", "shut down": "📴", "ventilation": "🌬️", "safety hazard": "☢️",
    "fixture": "🚿", "architect": "📐", "blueprint": "🗺️", "demolish": "💣",
}

# ---- Bài bổ sung: cụm từ hay gặp (từ sheet BUỔI 15–17, có thêm văn bản đọc) ----
COLLOCATIONS = [
    ("highly recommended / highly qualified", "Được đánh giá rất cao / cực kỳ đủ năng lực", "highly recommended. highly qualified"),
    ("strictly prohibited", "Bị nghiêm cấm", None),
    ("significantly / substantially increase", "Tăng đáng kể (decrease: giảm đáng kể)", "significantly increase. substantially increase"),
    ("currently available", "Hiện đang có sẵn", None),
    ("mutually beneficial", "Đôi bên cùng có lợi", None),
    ("promptly at 8:00 AM", "Đúng 8 giờ (promptly at + giờ)", "promptly at eight A M"),
    ("conveniently located", "Nằm ở vị trí thuận tiện", None),
    ("considerable effort", "Nỗ lực đáng kể", None),
    ("strict / stringent regulation", "Quy định nghiêm ngặt (policy / rule / regulation)", "strict regulation. stringent regulation"),
    ("meet the deadline", "Kịp thời hạn (meet the demands / needs: đáp ứng yêu cầu / nhu cầu)", None),
    ("conduct a survey", "Tiến hành khảo sát (research / study: nghiên cứu)", None),
    ("address an issue", "Giải quyết vấn đề (problem / concern: lo ngại)", None),
    ("implement a policy", "Thực thi chính sách (system / plan: hệ thống / kế hoạch)", None),
    ("extend an invitation", "Đưa ra lời mời (extend an offer: đưa ra lời đề nghị)", None),
    ("decline an offer", "Từ chối lời đề nghị", None),
    ("fill a vacancy", "Tuyển người vào vị trí còn trống (fill a position)", None),
    ("prove + adj", "Chứng tỏ (là)… — ví dụ: prove effective", "prove effective"),
    ("prove oneself", "Tự chứng tỏ (prove + đại từ phản thân)", None),
    ("prior to + N / V-ing", "Trước (khi)…", "prior to"),
    ("be eligible for + N / to V", "Đủ điều kiện (cho / để…)", "be eligible for"),
    ("at reasonable prices", "Với giá cả hợp lý (at reasonable rates)", None),
]

# ---- Bài bổ sung: liên từ / giới từ (cột G–H của BUỔI 15–18) ----
# Tài liệu gốc chỉ có danh sách từ; nghĩa, phiên âm và phân loại là phần bổ sung.
# Loại: C = liên từ (+ mệnh đề), P = giới từ (+ N / V-ing), A = trạng từ liên kết
CONNECTORS = [
    ("As soon as", "/əz ˈsuːn əz/", "Ngay khi", "C"),
    ("Whether", "/ˈweð.ɚ/", "Liệu… (hay không); dù… hay", "C"),
    ("Before", "/bɪˈfɔːr/", "Trước khi", "CP"),
    ("After", "/ˈæf.tɚ/", "Sau khi", "CP"),
    ("As", "/æz/", "Khi; vì; như là", "CP"),
    ("When", "/wen/", "Khi", "C"),
    ("While", "/waɪl/", "Trong khi; trong lúc", "C"),
    ("Once", "/wʌns/", "Một khi (Once + hiện tại, mệnh đề chính tương lai)", "C"),
    ("Since", "/sɪns/", "Kể từ khi; bởi vì", "CP"),
    ("During", "/ˈdʊr.ɪŋ/", "Trong suốt (khoảng thời gian)", "P"),
    ("Following", "/ˈfɑː.loʊ.ɪŋ/", "Sau; tiếp theo", "P"),
    ("Prior to", "/ˈpraɪ.ɚ tuː/", "Trước (khi)", "P"),
    ("Because", "/bɪˈkɑːz/", "Bởi vì", "C"),
    ("For", "/fɔːr/", "Cho; trong (bao lâu); vì", "CP"),
    ("Due to", "/ˈduː tuː/", "Do; bởi vì", "P"),
    ("Owing to", "/ˈoʊ.ɪŋ tuː/", "Do; bởi vì", "P"),
    ("On account of", "/ɑːn əˈkaʊnt əv/", "Do; vì lý do", "P"),
    ("Thanks to", "/ˈθæŋks tuː/", "Nhờ có", "P"),
    ("Though", "/ðoʊ/", "Mặc dù", "C"),
    ("Although", "/ɑːlˈðoʊ/", "Mặc dù", "C"),
    ("Even though", "/ˈiː.vən ðoʊ/", "Mặc dù (nhấn mạnh)", "C"),
    ("Even if", "/ˈiː.vən ɪf/", "Ngay cả khi; cho dù", "C"),
    ("Whereas", "/werˈæz/", "Trong khi (đối lập)", "C"),
    ("In spite of", "/ɪn ˈspaɪt əv/", "Mặc dù; bất chấp", "P"),
    ("Despite", "/dɪˈspaɪt/", "Mặc dù; bất chấp", "P"),
    ("Notwithstanding", "/ˌnɑːt.wɪθˈstæn.dɪŋ/", "Bất chấp; tuy vậy", "P"),
    ("Regardless of", "/rɪˈɡɑːrd.ləs əv/", "Bất kể", "P"),
    ("If", "/ɪf/", "Nếu", "C"),
    ("Only if", "/ˈoʊn.li ɪf/", "Chỉ khi", "C"),
    ("As long as", "/əz ˈlɑːŋ əz/", "Miễn là", "C"),
    ("Provided that", "/prəˈvaɪ.dɪd ðæt/", "Với điều kiện là; miễn là", "C"),
    ("Providing that", "/prəˈvaɪ.dɪŋ ðæt/", "Với điều kiện là; miễn là", "C"),
    ("On the condition that", "/ɑːn ðə kənˈdɪʃ.ən ðæt/", "Với điều kiện là", "C"),
    ("Supposing", "/səˈpoʊ.zɪŋ/", "Giả sử", "C"),
    ("In case", "/ɪn ˈkeɪs/", "Phòng khi", "C"),
    ("In case of", "/ɪn ˈkeɪs əv/", "Trong trường hợp (có)", "P"),
    ("In the event that", "/ɪn ði ɪˈvent ðæt/", "Trong trường hợp mà", "C"),
    ("But for", "/ˈbʌt fɔːr/", "Nếu không có; nếu không vì", "P"),
    ("Without", "/wɪˈðaʊt/", "Không có; nếu không có", "P"),
    ("Except that", "/ɪkˈsept ðæt/", "Ngoại trừ việc", "C"),
    ("As if / As though", "/əz ˈɪf/ · /əz ˈðoʊ/", "Như thể là", "C"),
    ("And", "/ænd/", "Và", "C"),
    ("But", "/bʌt/", "Nhưng", "C"),
    ("Yet", "/jet/", "Nhưng; tuy thế", "C"),
    ("So", "/soʊ/", "Vì vậy; nên", "C"),
    ("As well as", "/əz ˈwel əz/", "Cũng như", "P"),
    ("In addition to", "/ɪn əˈdɪʃ.ən tuː/", "Ngoài… ra; thêm vào", "P"),
    ("Besides", "/bɪˈsaɪdz/", "Ngoài ra; bên cạnh đó", "PA"),
    ("Furthermore", "/ˈfɝː.ðɚ.mɔːr/", "Hơn nữa", "A"),
    ("Moreover", "/mɔːrˈoʊ.vɚ/", "Hơn nữa", "A"),
    ("However", "/haʊˈev.ɚ/", "Tuy nhiên", "A"),
    ("Nevertheless", "/ˌnev.ɚ.ðəˈles/", "Tuy nhiên; dù vậy", "A"),
    ("Nonetheless", "/ˌnʌn.ðəˈles/", "Tuy nhiên; dù vậy", "A"),
    ("Still", "/stɪl/", "Tuy vậy; vẫn", "A"),
    ("Otherwise", "/ˈʌð.ɚ.waɪz/", "Nếu không thì", "A"),
    ("Therefore", "/ˈðer.fɔːr/", "Do đó; vì vậy", "A"),
    ("Thus", "/ðʌs/", "Do đó; như vậy", "A"),
    ("Hence", "/hens/", "Do đó; vì thế", "A"),
    ("Concerning", "/kənˈsɝː.nɪŋ/", "Liên quan đến; về việc", "P"),
    ("Regarding", "/rɪˈɡɑːr.dɪŋ/", "Liên quan đến; về việc", "P"),
    ("Related to", "/rɪˈleɪ.t̬ɪd tuː/", "Liên quan đến", "P"),
    ("According to", "/əˈkɔːr.dɪŋ tuː/", "Theo như", "P"),
    ("Based on", "/ˈbeɪst ɑːn/", "Dựa trên", "P"),
]
KIND = {
    "C": "Liên từ · + mệnh đề",
    "P": "Giới từ · + N / V-ing",
    "A": "Trạng từ liên kết",
    "CP": "Liên từ / Giới từ",
    "PA": "Giới từ / Trạng từ liên kết",
}


def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def norm_ipa(ipa):
    ipa = ipa.strip()
    if not ipa:
        return ""
    if " / " in ipa:  # hai cách đọc: "/a/ / /b/"
        return " · ".join(p.strip() for p in ipa.split(" / "))
    core = ipa.strip("/").strip()
    return f"/{core}/"


def cap(s):
    return s[:1].upper() + s[1:] if s else s


def read_sheet(ws, idx):
    topic, common = [], []
    bucket = topic
    for row in ws.iter_rows(min_col=1, max_col=4, values_only=True):
        a, b, c, d = row
        text = clean(a)
        if "THÔNG THUỘC" in text.upper():
            bucket = common
            continue
        if "CỤM TỪ" in text.upper():
            bucket = None  # phần cụm từ được nhập tay ở COLLOCATIONS
            continue
        if bucket is None or not isinstance(a, (int, float)) or not clean(b) or not clean(d):
            continue
        entry = {"w": cap(clean(b)), "ipa": norm_ipa(clean(c)), "vi": cap(clean(d))}
        entry.update(FIXES.get((idx, clean(b).lower()), {}))
        key = entry["w"].lower()
        if key in EMOJI:
            entry["e"] = EMOJI[key]
        if key in SAY:
            entry["say"] = SAY[key]
        bucket.append(entry)
    return topic, common


def with_ids(prefix, words):
    return [{"id": f"{prefix}-{i + 1}", **w} for i, w in enumerate(words)]


def main():
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    lessons = []
    for idx, ws in enumerate(wb.worksheets):
        label, topic_en, topic_vi, emoji, accent = LESSONS[idx]
        topic, common = read_sheet(ws, idx)
        lid = f"s{idx + 1}"
        decks = [{"id": f"{lid}t", "name": "Từ chủ điểm", "words": with_ids(f"{lid}t", topic)}]
        if common:
            decks.append({"id": f"{lid}c", "name": "Từ thông thuộc", "words": with_ids(f"{lid}c", common)})
        lessons.append({"id": lid, "label": label, "topic": topic_en, "topicVi": topic_vi,
                        "emoji": emoji, "accent": accent, "decks": decks})
        print(f"{label:8} {topic_en:36} chủ điểm={len(topic):2} thông thuộc={len(common):2}")

    colloc = []
    for w, vi, say in COLLOCATIONS:
        e = {"w": w, "ipa": "", "vi": vi, "noType": True}
        if say:
            e["say"] = say
        colloc.append(e)
    lessons.append({"id": "x1", "label": "Bổ trợ 1", "topic": "Must-know Collocations",
                    "topicVi": "Cụm từ chắc chắn gặp trong bài thi", "emoji": "🎯", "accent": "us",
                    "bonus": True,
                    "decks": [{"id": "x1t", "name": "Cụm từ", "words": with_ids("x1t", colloc)}]})

    conn = []
    for w, ipa, vi, kind in CONNECTORS:
        e = {"w": w, "ipa": ipa, "vi": vi, "note": KIND[kind]}
        if " / " in w:
            e["say"] = w.replace(" / ", ". ")
        conn.append(e)
    lessons.append({"id": "x2", "label": "Bổ trợ 2", "topic": "Conjunctions & Prepositions",
                    "topicVi": "Liên từ & Giới từ", "emoji": "🔗", "accent": "us", "bonus": True,
                    "decks": [{"id": "x2t", "name": "Liên từ / Giới từ", "words": with_ids("x2t", conn)}]})

    total = sum(len(d["words"]) for l in lessons for d in l["decks"])
    missing = sorted({w["w"] for l in lessons if not l.get("bonus") for d in l["decks"]
                      for w in d["words"] if "e" not in w})
    print(f"Tổng: {len(lessons)} bài, {total} mục. Chưa có emoji riêng: {len(missing)}")
    print(", ".join(missing))

    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"lessons": lessons}, ensure_ascii=False, separators=(",", ":"))
    OUT.write_text("// Tệp sinh tự động bởi tools/build_data.py — đừng sửa tay.\n"
                   "// © 2026 Vườn từ của cô Vân Anh. Bảo lưu mọi quyền. Không sao chép, phân phối hoặc dùng lại "
                   "nội dung bài học khi chưa được phép (xem tệp LICENSE).\nwindow.VOCAB = "
                   + payload + ";\n", encoding="utf-8")
    print(f"Đã ghi {OUT} ({OUT.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
