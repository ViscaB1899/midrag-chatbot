## התקנת ספריות נדרשות
!pip install openai langchain langchain-community faiss-cpu google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client tiktoken python-docx pandas openpyxl

import os
import io
import time
import docx
import pandas as pd
from openai import OpenAI
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS
from langchain.memory import ConversationBufferMemory


# מחלקה לחיבור לגוגל דרייב והורדת קבצים
class DriveConnector:
    def __init__(self):
        auth.authenticate_user()
        self.drive_service = build('drive', 'v3')

    def get_file_content_by_name(self, file_name):
        results = self.drive_service.files().list(q=f"name='{file_name}'", fields="files(id)").execute()
        files = results.get('files', [])
        if not files:
            print(f"לא נמצא קובץ בשם: {file_name}")
            return None
        file_id = files[0]['id']
        request = self.drive_service.files().get_media(fileId=file_id)
        file_content = io.BytesIO()
        downloader = MediaIoBaseDownload(file_content, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        # החזרת התוכן כבייטים לקבצי Word ו-Excel ללא פענוח
        if file_name.endswith(('.docx', '.xlsx')):
            return file_content.getvalue()
        # אחרת מחזירים טקסט מפוענח
        return file_content.getvalue().decode('utf-8')


# מחלקת הצ'אטבוט המלאה
class ChatBot:
    def __init__(self, knowledge_files, api_key_file):
        self.drive_connector = DriveConnector()
        self.api_key = os.environ.get("OPENAI_API_KEY")
        self.client = OpenAI(api_key=self.api_key)
        self.embeddings = OpenAIEmbeddings(api_key=self.api_key)
        self.memory = ConversationBufferMemory()
        # תמיכה ברשימת קבצים
        self.knowledge_files = knowledge_files if isinstance(knowledge_files, list) else [knowledge_files]
        self.vectorstore = None

    def load_knowledge_base(self):
        all_chunks = []

        for file_name in self.knowledge_files:
            if file_name.endswith('.docx'):
                # טיפול בקובץ Word
                chunks = self.process_docx_file(file_name)
                print(f"קובץ Word {file_name} נטען עם {len(chunks)} צ'אנקים.")
            elif file_name.endswith('.xlsx'):
                # טיפול בקובץ Excel
                chunks = self.process_excel_file(file_name)
                print(f"קובץ Excel {file_name} נטען עם {len(chunks)} צ'אנקים.")
            else:
                # טיפול בקובץ טקסט רגיל כמו בקוד המקורי
                content = self.drive_connector.get_file_content_by_name(file_name)
                if not content:
                    print(f"הקובץ {file_name} לא נמצא או לא נטען כראוי.")
                    continue
                chunks = self.split_to_chunks(content)
                print(f"קובץ טקסט {file_name} נטען עם {len(chunks)} צ'אנקים.")

            all_chunks.extend(chunks)

        if not all_chunks:
            raise ValueError("לא הצלחנו לטעון אף קובץ מבסיס הידע.")

        self.vectorstore = FAISS.from_texts(all_chunks, self.embeddings)
        print(f"בסיס הידע נטען בהצלחה עם סה״כ {len(all_chunks)} צ'אנקים.")

    def process_docx_file(self, file_name):
        """מעבד קובץ Word ומחלק אותו לצ'אנקים"""
        file_bytes = io.BytesIO(self.drive_connector.get_file_content_by_name(file_name))
        doc = docx.Document(file_bytes)
        text_content = "\n\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        return self.split_to_chunks(text_content)

    def process_excel_file(self, file_name):
        """מעבד קובץ Excel ומחלק אותו לצ'אנקים"""
        file_bytes = io.BytesIO(self.drive_connector.get_file_content_by_name(file_name))
        # קריאת כל הגיליונות בקובץ Excel
        excel_file = pd.ExcelFile(file_bytes)
        all_texts = []

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            # שימוש בכותרות עמודות
            headers = df.columns.tolist()

            # עיבוד כל שורה בגיליון
            for idx, row in df.iterrows():
                row_text = f"גיליון: {sheet_name}\n"
                for col in headers:
                    if pd.notna(row[col]) and str(row[col]).strip():
                        row_text += f"{col}: {row[col]}\n"

                if len(row_text) > 50:  # רק אם יש מספיק מידע בשורה
                    all_texts.append(row_text)

        # פיצול טקסטים ארוכים לצ'אנקים
        chunks = []
        for text in all_texts:
            if len(text) > 1500:
                chunks.extend(self.split_to_chunks(text))
            else:
                chunks.append(text)

        return chunks

    def split_to_chunks(self, text, max_chunk_size=1500, overlap=100):
        sentences = text.split('. ')
        chunks = []
        current_chunk = ""
        for sentence in sentences:
            if len(current_chunk + sentence) <= max_chunk_size:
                current_chunk += sentence + '. '
            else:
                chunks.append(current_chunk.strip())
                current_chunk = sentence + '. '
        if current_chunk:
            chunks.append(current_chunk.strip())
        return chunks

    def get_relevant_chunks(self, question, k=3):
        docs = self.vectorstore.similarity_search(question, k=k)
        return [doc.page_content for doc in docs]

    def ask(self, question):
        context = '\n\n---\n\n'.join(self.get_relevant_chunks(question))
        history = self.memory.load_memory_variables({}).get('history', '')

        prompt = f"""
        אתה עוזר מקצועי, מדויק וברור, המסייע לנציגי שירות במחלקת קליטת עסקים באתר מידרג.
        התשובה צריכה להיות מפורטת, ברורה, ומדויקת על סמך המידע המוצג בלבד.
        אם התשובה לא קיימת במידע, ציין זאת מפורשות. תן תשובות של מקסימום 3 שורות

        היסטוריית השיחה הקודמת:
        {history}

        מידע רלוונטי לשאלה:
        {context}

        שאלה:
        {question}
        """

        response = self.client.chat.completions.create(
            model="gpt-4-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=1000
        )

        answer = response.choices[0].message.content.strip()
        self.memory.save_context({"input": question}, {"output": answer})

        return answer

    def chat(self):
        print("הצ'אטבוט פעיל ומוכן לשאלותיך. רשום 'יציאה' כדי לסיים את השיחה.")
        while True:
            question = input("🔸 השאלה שלך: ")
            if question.lower() == 'יציאה':
                print("הצ'אט הסתיים, תודה ולהתראות!")
                break
            start_time = time.time()
            answer = self.ask(question)
            elapsed_time = time.time() - start_time
            print(f"\n🔹 תשובה: {answer}\n⏱️ זמן תגובה: {elapsed_time:.2f} שניות\n")


# הפעלת הצ'אטבוט
knowledge_files = ["הסבר על תהליך ההצטרפות 3.txt"]
bot = ChatBot(knowledge_files, "openai_key.txt")
bot.load_knowledge_base()
bot.chat()
