from flask import Flask, request, jsonify
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

@app.route("/webhook", methods=["POST"])
def webhook():
    req = request.json
    command = req['request']['original_utterance']

    # Пример: "Добавь в таблицу Иван заказ 123 сумма 4500"
    # Простая обработка (можно улучшать с NLP)
    try:
        name = command.split(" ")[3]
        order = command.split("заказ")[1].split("сумма")[0].strip()
        amount = command.split("сумма")[1].strip()

        # Открытие Excel-файла
        wb = load_workbook("orders.xlsx")
        sheet = wb.active
        sheet.append([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), name, order, amount])
        wb.save("orders.xlsx")

        response_text = f"Добавил: {name}, заказ {order}, сумма {amount}."
    except Exception as e:
        response_text = f"Произошла ошибка: {str(e)}"

    return jsonify({
        "response": {
            "text": response_text,
            "end_session": True
        },
        "version": req['version']
    })

if __name__ == "__main__":
    app.run()