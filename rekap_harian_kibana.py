import glob
import os
import zipfile
import pandas as pd


def clean_num(val):
  """Fungsi pembersih angka dari format ribuan ber-koma/string ke float/int"""
  if pd.isna(val):
    return 0
  if isinstance(val, (int, float)):
    return val
  return float(str(val).replace(',', ''))


def process_dashboard3(zip_file_path):
  """Memproses Dashboard 3: CMD Performance Dashboard menjadi 1 BARIS PER TANGGAL (Pivoted Header)

  Menghasilkan File Excel & Teks Final Summary khusus Modul Administration.
  """
  print(f'\n==========================================')
  print(
      '📌 DASHBOARD 3: EXECUTIVE SUMMARY (1 BARIS PER TANGGAL)'
      f' ({os.path.basename(zip_file_path)})'
  )
  print(f'==========================================')

  extract_folder = 'extracted_d3'
  with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_folder)

  all_csvs = glob.glob(f'{extract_folder}/**/*.csv', recursive=True)

  data_by_date = {}
  for csv_path in all_csvs:
    parts = os.path.normpath(csv_path).split(os.sep)
    date_folder = None
    for p in parts:
      if p.isdigit():
        date_folder = int(p)
        break

    if date_folder is not None:
      if date_folder not in data_by_date:
        data_by_date[date_folder] = {}

      filename = os.path.basename(csv_path).lower()
      if 'admin' in filename:
        domain = 'Administration'
      elif 'config' in filename:
        domain = 'Config'
      elif 'user' in filename:
        domain = 'User Management'
      elif 'interface' in filename:
        domain = 'Interface'
      elif 'integration' in filename:
        domain = 'Integration'
      elif 'logging' in filename:
        domain = 'Logging'
      elif 'dashboard' in filename:
        domain = 'Dashboard'

      data_by_date[date_folder][domain] = csv_path

  all_domains = [
      'Administration',
      'Config',
      'User Management',
      'Interface',
      'Integration',
      'Logging',
      'Dashboard',
  ]
  rows_d3 = []

  # Menampung nilai harian khusus Modul Administration
  admin_avg_list = []
  admin_p90_list = []

  for day in sorted(data_by_date.keys()):
    date_str = f'{day:02d} Sep 2026'
    day_files = data_by_date[day]

    row_data = {'Tanggal': date_str}

    for dom in all_domains:
      if dom in day_files:
        df = pd.read_csv(day_files[dom])
        avg_col = [c for c in df.columns if 'average' in c.lower()][0]
        p90_col = [c for c in df.columns if '90th' in c.lower()][0]
        rec_col = [c for c in df.columns if 'records' in c.lower()][0]

        df['Avg'] = df[avg_col].apply(clean_num)
        df['P90'] = df[p90_col].apply(clean_num)
        df['Rec'] = df[rec_col].apply(clean_num)

        tot_rec = df['Rec'].sum()
        if tot_rec > 0:
          w_avg = (df['Avg'] * df['Rec']).sum() / tot_rec
          w_p90 = (df['P90'] * df['Rec']).sum() / tot_rec
        else:
          w_avg, w_p90 = 0, 0

        avg_val = int(round(w_avg))
        p90_val = int(round(w_p90))

        row_data[f'{dom} Avg (ms)'] = avg_val
        row_data[f'{dom} P90 (ms)'] = p90_val

        # Simpan nilai khusus modul Administration
        if dom == 'Administration':
          if avg_val > 0:
            admin_avg_list.append(avg_val)
          if p90_val > 0:
            admin_p90_list.append(p90_val)
      else:
        row_data[f'{dom} Avg (ms)'] = 0
        row_data[f'{dom} P90 (ms)'] = 0

    rows_d3.append(row_data)

  df_out3 = pd.DataFrame(rows_d3)
  excel_d3 = 'Hasil_Rekap_Dashboard_3_Performance.xlsx'
  df_out3.to_excel(excel_d3, index=False)

  # Hitung Rata-rata Sederhana Khusus Modul Administration (Abaikan nilai 0)
  avg_admin_all = (
      int(round(sum(admin_avg_list) / len(admin_avg_list)))
      if admin_avg_list
      else 0
  )
  p90_admin_all = (
      int(round(sum(admin_p90_list) / len(admin_p90_list)))
      if admin_p90_list
      else 0
  )

  summary_d3_text = (
      'CMD Performance Dashboard: Pemantauan latensi pada Modul Administration (sebagai modul utama paling aktif)'
      f' mencatatkan Rata-rata Latency sebesar {avg_admin_all} ms dengan batas'
      ' kenyamanan mayoritas pengguna (90th Percentile / P90) berada pada angka'
      f' {p90_admin_all} ms.'
  )

  print(f'✅ File Excel Dashboard 3 berhasil dibuat: {excel_d3}')
  print(f'\n📝 FINAL SUMMARY D3:\n{summary_d3_text}\n')
  return summary_d3_text


def process_dashboard2_summary(zip_file_path):
  """Memproses Dashboard 2: CMD Monitoring Dashboard (1 Baris per Tanggal)

  Termasuk Total Trx, Total Success, Total Error, Success Rate %, Error Rate %,
  Avg Response Time, serta mencetak Final Summary Naratif otomatis.
  """
  print(f'\n==========================================')
  print(
      '📌 DASHBOARD 2: RINGKASAN HARIAN / SUMMARY'
      f' ({os.path.basename(zip_file_path)})'
  )
  print(f'==========================================')

  extract_folder = 'extracted_d2'
  with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_folder)

  all_csvs = glob.glob(f'{extract_folder}/**/*.csv', recursive=True)

  data_by_date = {}
  for csv_path in all_csvs:
    parts = os.path.normpath(csv_path).split(os.sep)
    date_folder = None
    for p in parts:
      if p.isdigit():
        date_folder = int(p)
        break

    if date_folder is not None:
      if date_folder not in data_by_date:
        data_by_date[date_folder] = {}

      fn = os.path.basename(csv_path).lower()
      if 'total transaction' in fn:
        data_by_date[date_folder]['tot'] = csv_path
      elif 'total trx success' in fn:
        data_by_date[date_folder]['succ'] = csv_path
      elif 'total trx error' in fn:
        data_by_date[date_folder]['err'] = csv_path
      elif 'response time' in fn:
        data_by_date[date_folder]['rt'] = csv_path

  rows_d2 = []

  grand_tot_trx = 0
  grand_succ_trx = 0
  grand_err_trx = 0
  total_rt_weighted_sum = 0

  for day in sorted(data_by_date.keys()):
    date_str = f'{day:02d} Sep 2026'
    d_files = data_by_date[day]

    tot_trx, succ_trx, err_trx, w_rt = 0, 0, 0, 0

    if 'tot' in d_files:
      df_tot = pd.read_csv(d_files['tot'])
      c_tot_rec = [
          c
          for c in df_tot.columns
          if 'count' in c.lower() or 'records' in c.lower()
      ][0]
      df_tot['Tot_clean'] = df_tot[c_tot_rec].apply(clean_num)
      tot_trx = int(df_tot['Tot_clean'].sum())
      svc_col_tot = df_tot.columns[0]

      if 'succ' in d_files:
        df_succ = pd.read_csv(d_files['succ'])
        c_succ_rec = [
            c
            for c in df_succ.columns
            if 'count' in c.lower() or 'records' in c.lower()
        ][0]
        df_succ['Succ_clean'] = df_succ[c_succ_rec].apply(clean_num)
        succ_trx = int(df_succ['Succ_clean'].sum())

      if 'err' in d_files:
        df_err = pd.read_csv(d_files['err'])
        c_err_rec = [
            c
            for c in df_err.columns
            if 'count' in c.lower() or 'records' in c.lower()
        ][0]
        df_err['Err_clean'] = df_err[c_err_rec].apply(clean_num)
        err_trx = int(df_err['Err_clean'].sum())
      else:
        err_trx = tot_trx - succ_trx if tot_trx >= succ_trx else 0

      if 'rt' in d_files:
        df_rt = pd.read_csv(d_files['rt'])
        c_rt_val = [
            c
            for c in df_rt.columns
            if 'average' in c.lower() or 'time' in c.lower()
        ][0]
        svc_col_rt = df_rt.columns[0]
        df_rt['Rt_clean'] = df_rt[c_rt_val].apply(clean_num)

        df_merged = pd.merge(
            df_tot, df_rt, left_on=svc_col_tot, right_on=svc_col_rt, how='inner'
        )
        if tot_trx > 0:
          w_rt = (
              df_merged['Tot_clean'] * df_merged['Rt_clean']
          ).sum() / tot_trx
        else:
          w_rt = 0

    sr = (succ_trx / tot_trx * 100) if tot_trx > 0 else 0.0
    er = (err_trx / tot_trx * 100) if tot_trx > 0 else 0.0

    rt_val = int(round(w_rt))
    sr_str = f'{sr:.2f}%'
    er_str = f'{er:.2f}%'

    grand_tot_trx += tot_trx
    grand_succ_trx += succ_trx
    grand_err_trx += err_trx
    total_rt_weighted_sum += w_rt * tot_trx

    rows_d2.append({
        'Tanggal': date_str,
        'Total Transaksi': tot_trx,
        'Total Success': succ_trx,
        'Total Error': err_trx,
        'Success Rate (%)': sr_str,
        'Error Rate (%)': er_str,
        'Avg Response Time (ms)': rt_val,
    })

  df_out2 = pd.DataFrame(rows_d2)
  excel_d2 = 'Hasil_Rekap_Dashboard_2_Monitoring.xlsx'
  df_out2.to_excel(excel_d2, index=False)

  # Hitung Nilai Akumulasi Keseluruhan Periode
  overall_sr = (
      (grand_succ_trx / grand_tot_trx * 100) if grand_tot_trx > 0 else 0.0
  )
  overall_er = (
      (grand_err_trx / grand_tot_trx * 100) if grand_tot_trx > 0 else 0.0
  )
  overall_rt = (
      int(round(total_rt_weighted_sum / grand_tot_trx))
      if grand_tot_trx > 0
      else 0
  )

  summary_d2_text = (
      'CMD Monitoring Dashboard: Memantau service level harian dari seluruh'
      ' microservices dengan total volume transaksi mencapai'
      f' {grand_tot_trx:,} transaksi ({grand_succ_trx:,} transaksi sukses dan'
      f' {grand_err_trx:,} transaksi error), menghasilkan rata-rata Success'
      f' Rate sebesar {overall_sr:.2f}%, Error Rate {overall_er:.2f}%, serta'
      ' Weighted Average Response Time terjaga cepat di angka'
      f' {overall_rt} ms.'
  )

  print(f'✅ File Excel Dashboard 2 berhasil dibuat: {excel_d2}')
  print(f'\n📝 FINAL SUMMARY D2:\n{summary_d2_text}\n')
  return summary_d2_text


if __name__ == '__main__':
  zip_files = glob.glob('*.zip')

  if not zip_files:
    print('❌ Tidak ditemukan file .zip di folder ini!')
  else:
    for z in zip_files:
      fn = os.path.basename(z).upper()

      if 'D2' in fn or 'MONITORING' in fn or 'CSV 2' in fn:
        process_dashboard2_summary(z)
      else:
        process_dashboard3(z)