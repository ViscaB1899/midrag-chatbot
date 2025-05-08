from flask import Flask, request, jsonify
from chatbot import ChatBot
import threading
import os
import time

# יצירת מופע הצ'אטבוט
knowledge_files = ["הסבר על תהליך ההצטרפות 3.txt"]
bot = ChatBot(knowledge_files, None)
knowledge_loaded = False

# פונקציה לטעינת בסיס הידע
def load_knowledge():
    global knowledge_loaded
    try:
        bot.load_knowledge_base()
        knowledge_loaded = True
        print("✅ בסיס הידע נטען בהצלחה.")
    except Exception as e:
        print(f"❌ שגיאה בטעינת בסיס הידע: {e}")

# הרצת טעינת בסיס הידע בשרשור נפרד
threading.Thread(target=load_knowledge).start()

# הגדרת Flask
app = Flask(__name__)

@app.route("/", methods=["POST"])
def webhook():
    data = request.get_json()
    try:
        question = data.get("message", {}).get("text", "")
        if not question:
            return jsonify({"text": "לא הוזנה שאלה תקינה"}), 400
        if not knowledge_loaded:
            return jsonify({"text": "השרת עדיין טוען את בסיס הידע. נסה שוב בעוד מספר שניות."}), 503
        answer = bot.ask(question)
        return jsonify({"text": answer})
    except Exception as e:
        print(f"❌ שגיאה בטיפול בשאלה: {e}")
        return jsonify({"text": "שגיאה פנימית בשרת"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
