import pyttsx3

def speak(text):
    """
    Speak normal English sentences reliably.
    Creates a fresh engine every call to avoid pyttsx3 deadlock.
    """
    if not isinstance(text, str):
        text = str(text)

    text = text.strip()
    if not text:
        return

    engine = pyttsx3.init()
    engine.setProperty("rate", 170)
    engine.setProperty("volume", 1.0)

    # Optional: try to use Microsoft David
    for v in engine.getProperty("voices"):
        if "DAVID" in v.id.upper() and "EN-US" in v.id.upper():
            engine.setProperty("voice", v.id)
            break

    engine.say(text)
    engine.runAndWait()
    engine.stop()