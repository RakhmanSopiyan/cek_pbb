import os
import re
import requests
import pandas as pd
from flask import Flask, request, render_template, send_file
from bs4 import BeautifulSoup
import concurrent.futures

app = Flask(__name__)

URL = "http://bogorkab.net/cekpbb/cekpbb/op?nop_kd={}"
TAHUN_TARGET = ["2021", "2022", "2023", "2024", "2025"]

def extract_target_years(data_pbb, tahun_list):
    data_filtered = []
    for row in data_pbb:
        for year in tahun_list:
            if any(re.search(rf'\b{year}\b', cell) for cell in row):
                data_filtered.append((year, row))
                break
    return data_filtered

def cek_pbb(nop):
    try:
        response = requests.get(URL.format(nop), timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "lxml")
            table = soup.find("table")

            if table:
                rows = table.find_all("tr")
                data_pbb = []
                for row in rows:
                    cols = row.find_all("td")
                    cols = [col.text.strip() for col in cols]
                    if cols:
                        data_pbb.append(cols)

                data_filtered = extract_target_years(data_pbb, TAHUN_TARGET)
                if data_filtered:
                    results = []
                    for tahun, row in data_filtered:
                        hasil = " - ".join(row)
                        results.append([nop, "Berhasil", tahun, hasil])
                    return results
                else:
                    return [[nop, "Tidak Ada Data", "-", "-"]]
            else:
                return [[nop, "Data Tidak Ditemukan", "-", "-"]]
        else:
            return [[nop, "Gagal", "-", "Tidak dapat diakses"]]
    except Exception as e:
        return [[nop, "Error", "-", str(e)]]


@app.route("/", methods=["GET", "POST"])
def index():
    hasil_data = []
    if request.method == "POST":
        file = request.files["file"]
        if file:
            content = file.read().decode("utf-8")
            list_nop = [line.strip() for line in content.splitlines() if line.strip()]

            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                results = executor.map(cek_pbb, list_nop)
                for res in results:
                    hasil_data.extend(res)

            df = pd.DataFrame(hasil_data, columns=["NOP", "Status", "Tahun", "Detail"])
            output_path = os.path.join("hasil.xlsx")
            df.to_excel(output_path, index=False, engine="openpyxl")
            return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
