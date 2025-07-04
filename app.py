from flask import Flask, request, jsonify
from openpyxl import load_workbook
from datetime import datetime
import os

from openpyxl.workbook import Workbook

app = Flask(__name__)

EXCEL_PATH = "orders.xlsx"


@app.route("/", methods=["GET"])
def index():
    return "Навык Алисы работает!"


@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        req = request.get_json()

        # Безопасно достаём поля из запроса
        command = req.get("request", {}).get("original_utterance", "").lower()

        if not command:
            return jsonify({
                "response": {
                    "text": "Я не услышала команду. Попробуй ещё раз.",
                    "end_session": True
                },
                "version": req.get("version", "1.0")
            })

        # Пример: "добавь в таблицу Иван заказ 222 сумма 3000"
        try:
            name = command.split("таблицу")[1].split("заказ")[0].strip()
            order = command.split("заказ")[1].split("сумма")[0].strip()
            amount = command.split("сумма")[1].strip()
        except Exception as e:
            return jsonify({
                "response": {
                    "text": "Не смог разобрать данные. Скажи: добавь в таблицу Иван заказ 123 сумма 4000.",
                    "end_session": True
                },
                "version": req.get("version", "1.0")
            })
        if not os.path.exists(EXCEL_PATH):
            wb = Workbook()
            ws = wb.active
            ws.append(["Дата", "Имя", "Заказ", "Сумма"])  # заголовки
            wb.save(EXCEL_PATH)
        wb = load_workbook(EXCEL_PATH)
        sheet = wb.active
        sheet.append([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), name, order, amount])
        wb.save(EXCEL_PATH)
        print("Сохраняем в файл:", os.path.abspath(EXCEL_PATH), flush=True)
        return jsonify({
            "response": {
                "text": f"Добавил: {name}, заказ {order}, сумма {amount}.",
                "end_session": True
            },
            "version": req.get("version", "1.0")
        })

    except Exception as e:
        return jsonify({
            "response": {
                "text": f"Внутренняя ошибка: {str(e)}",
                "end_session": True
            },
            "version": "1.0"
        }), 200  # ❗ 200, чтобы Алиса не считала это ошибкой


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)