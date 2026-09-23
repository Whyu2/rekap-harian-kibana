# 📊 Kibana Daily Report Automation (`rekap_harian_kibana.py`)

Script Python otomatis untuk mengekstrak, mengolah, dan merakapitulasi ekspor file `.zip` berisi data CSV dari **Kibana Dashboard** ke dalam file laporan berformat **Microsoft Excel (`.xlsx`)**.

Script ini dirancang khusus untuk memenuhi kebutuhan pelaporan harian **Cash Management Dashboard (CMD)** pada dua area pemantauan utama:
1. **Dashboard 2 (CMD Monitoring Dashboard)**: Service Level, Volume Transaksi, Success/Error Rate, dan Response Time[cite: 8].
2. **Dashboard 3 (CMD Performance Dashboard)**: Latency & 90th Percentile per Modul Utama[cite: 9, 10].

---

## 📁 Struktur Input File & Konvensi Penamaan

Agar script dapat mendeteksi dan mengolah file secara otomatis tanpa risiko tertukar, pastikan file `.zip` ekspor Kibana diletakkan di folder yang sama dengan script dan diberi penamaan sesuai aturan berikut:

| Laporan Target | Contoh Nama File Zip | Deskripsi Isi File | Output File Excel |
| :--- | :--- | :--- | :--- |
| **Dashboard 2** | `D2_22.zip`<br>`CMD_D2_22Sep2026.zip`<br>`CSV 2.zip` | Folder harian berisi CSV *Total Transaction*, *Total Trx Success*, *Total Trx Error*, dan *Response Time* per service[cite: 8]. | `Hasil_Rekap_Dashboard_2_Monitoring.xlsx`[cite: 8] |
| **Dashboard 3** | `D3_22.zip`<br>`CMD_D3_22Sep2026.zip`<br>`22.zip` | Folder harian berisi CSV *Top API Call* per modul (`Administration`, `Config`, `User Management`, dll.)[cite: 9, 10]. | `Hasil_Rekap_Dashboard_3_Performance.xlsx`[cite: 9] |

---

## 🚀 Fitur Utama & Logika Kalkulasi

### 📊 1. Dashboard 2 (CMD Monitoring Dashboard)
* **Format Output**: **1 Baris per Tanggal** (Merangkum gabungan seluruh *microservices*)[cite: 8].
* **Skema Kolom**:
  `Tanggal` | `Total Transaksi` | `Total Success` | `Total Error` | `Success Rate (%)` | `Error Rate (%)` | `Avg Response Time (ms)`
* **Logika Perhitungan**:
  * **Success Rate (%)**: $\frac{\text{Total Success}}{\text{Total Transaksi}} \times 100\%$[cite: 8]
  * **Error Rate (%)**: $\frac{\text{Total Error}}{\text{Total Transaksi}} \times 100\%$[cite: 8]
  * **Avg Response Time**: Di-kalkulasi menggunakan *Weighted Average* (Rata-rata tertimbang berdasarkan volume transaksi masing-masing *service*)[cite: 8].

### 📈 2. Dashboard 3 (CMD Performance Dashboard)
* **Format Output**: **1 Baris per Tanggal (Pivoted / Mendatar)**[cite: 9].
* **Skema Kolom**:
  `Tanggal` serta metrik `Avg Latency (ms)` dan `90th Percentile (ms)` untuk **7 Modul Utama**:
  1. *Administration*[cite: 10]
  2. *Config*[cite: 10]
  3. *User Management*[cite: 10]
  4. *Interface*[cite: 10]
  5. *Integration*[cite: 10]
  6. *Logging*[cite: 10]
  7. *Dashboard*[cite: 10]
* **Logika Perhitungan**:
  * **Weighted Avg Latency**: $\frac{\sum (\text{Avg Latency API} \times \text{Count Records})}{\sum \text{Count Records}}$[cite: 9, 10]
  * **Weighted P90 Latency**: $\frac{\sum (\text{P90 Latency API} \times \text{Count Records})}{\sum \text{Count Records}}$[cite: 9, 10]

---

## 🛠️ Prasyarat & Instalasi

Pastikan komputer Anda telah terinstal **Python 3.x** dan pustaka **pandas** & **openpyxl**.

Jalankan perintah ini di Terminal / Command Prompt jika library belum terinstal:
```bash
pip install pandas openpyxl