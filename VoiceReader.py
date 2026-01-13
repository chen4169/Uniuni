import pyttsx3
import re

# 初始化引擎（只初始化一次）
_engine = pyttsx3.init()
_voices = _engine.getProperty("voices")

# 找到 David 英文男声
DAVID_VOICE_ID = None
for v in _voices:
    if "DAVID" in v.id.upper() and "EN-US" in v.id.upper():
        DAVID_VOICE_ID = v.id
        break

if DAVID_VOICE_ID is None:
    raise RuntimeError("Microsoft David English voice not found")

_engine.setProperty("voice", DAVID_VOICE_ID)
_engine.setProperty("rate", 170)
_engine.setProperty("volume", 1.0)


def spell_number(number):
    """
    把数字转成逐位英文读法
    202511 -> 'two zero two five one one'
    """
    return " ".join(list(str(number)))


def _replace_numbers(text):
    """
    自动把文本中的数字替换成 spell_number 形式
    """
    def replacer(match):
        return spell_number(match.group())

    return re.sub(r"\d+", replacer, text)


def speak(text):
    """
    使用 David 英文男声播报文本
    - 自动处理数字
    - 同步阻塞（说完才继续）
    """
    if not isinstance(text, str):
        text = str(text)

    processed_text = _replace_numbers(text)
    _engine.say(processed_text)
    _engine.runAndWait()