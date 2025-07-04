from flask import Flask, request, jsonify
from openpyxl import load_workbook
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_PATH = "orders.xlsx"

@app.route("/")
def index():
    return "Навык Алисы работает!"

@app.route("/webhook", methods=["POST"])
def webhook():
    req = request.json
    command = req['request']['original_utterance']

    try:
        name = command.split(" ")[3]
        order = command.split("заказ")[1].split("сумма")[0].strip()
        amount = command.split("сумма")[1].strip()

        wb = load_workbook(EXCEL_PATH)
        sheet = wb.active
        sheet.append([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), name, order, amount])
        wb.save(EXCEL_PATH)

        response_text = f"Добавил: {name}, заказ {order}, сумма {amount}."
    except Exception as e:
        response_text = f"Ошибка при обработке: {str(e)}"

    return jsonify({
        "response": {
            "text": response_text,
            "end_session": True
        },
        "version": req.get("version", "1.0")
    })

# Render требует указания порта через переменную окружения
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
