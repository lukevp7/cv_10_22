import os
import history_forecast, managers_flash


for file in os.listdir("reports_excel/"):
    if file.endswith(".xlsx"):
        os.remove(("reports_excel/"+file))

for file in os.listdir("reports_pdf/"):
        if file.endswith(".pdf"):
            filename = file


# history_forecast.start(filename)
# managers_flash.start(filename)


for file in os.listdir("reports_pdf/"):
    if file.endswith(".pdf"):
        os.remove(("reports_pdf/"+file))
