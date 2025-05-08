from flask import Flask, request, jsonify
from chatbot import ChatBot
import threading
import os

knowledge_files = ["הסבר על תהליך ההצטרפות 3.txt"]
bot = ChatBot(knowledge_files, None)
threading.Thread(target=bot.load_knowledge_base).start()

app = Flask(__name__)

@app.route("/", methods=["POST"])
def webhook():
    data = request.get_json()
    try:
        question = data.get("message", {}).get("text", "")
        if not question:
            return jsonify({"text": "לא הוזנה שאלה תקינה"}), 400
        answer = bot.ask(question)
        return jsonify({"text": answer})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"text": "שגיאה פנימית בשרת"}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
