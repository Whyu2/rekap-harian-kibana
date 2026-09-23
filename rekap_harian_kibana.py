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
  """Memproses Dashboard 3: CMD Performance Dashboard menjadi 1 BARIS PER TANGGAL (Pivoted Header)"""
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

        row_data[f'{dom} Avg (ms)'] = int(round(w_avg))
        row_data[f'{dom} P90 (ms)'] = int(round(w_p90))
      else:
        row_data[f'{dom} Avg (ms)'] = 0
        row_data[f'{dom} P90 (ms)'] = 0

    rows_d3.append(row_data)

  df_out3 = pd.DataFrame(rows_d3)
  excel_d3 = 'Hasil_Rekap_Dashboard_3_Performance.xlsx'
  df_out3.to_excel(excel_d3, index=False)
  print(f'✅ File Excel Dashboard 3 berhasil dibuat: {excel_d3}')


def process_dashboard2_summary(zip_file_path):
  """Memproses Dashboard 2: CMD Monitoring Dashboard (1 Baris per Tanggal)

  Termasuk Total Trx, Total Success, Total Error, Success Rate %, Error Rate %,
  dan Avg Response Time
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

    rows_d2.append({
        'Tanggal': date_str,
        'Total Transaksi': tot_trx,
        'Total Success': succ_trx,
        'Total Error': err_trx,
        'Success Rate (%)': sr_str,
        'Error Rate (%)': er_str,
        'Avg Response Time (ms)': rt_val,
    })

  # Export ke Excel Khusus Dashboard 2
  df_out2 = pd.DataFrame(rows_d2)
  excel_d2 = 'Hasil_Rekap_Dashboard_2_Monitoring.xlsx'
  df_out2.to_excel(excel_d2, index=False)

  # Tampilkan Ringkasan di Terminal
  print(df_out2.to_string(index=False))
  print(f'\n✅ File Excel Dashboard 2 berhasil dibuat: {excel_d2}')


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