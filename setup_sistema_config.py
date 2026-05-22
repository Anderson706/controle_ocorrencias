import oracledb, os

dsn = oracledb.makedsn('usqasap023-scan.phx-dc.dhl.com', 1521, service_name='SECPANEL')
conn = oracledb.connect(user='SECPANEL', password='SEC003q2w3e4r2026', dsn=dsn)
cur = conn.cursor()

exe_path = os.path.join(os.path.dirname(__file__), 'dist', 'CCTV_ControlPanel.exe')
print(f'EXE: {exe_path}  ({os.path.getsize(exe_path)/1_048_576:.1f} MB)')

with open(exe_path, 'rb') as f:
    exe_bytes = f.read()

cur.execute(
    "UPDATE SISTEMA_CONFIG SET EXE_BLOB = :blob, VERSAO_EXIGIDA = '3.0'",
    {'blob': exe_bytes}
)
conn.commit()

cur.execute("SELECT VERSAO_EXIGIDA, CASE WHEN EXE_BLOB IS NOT NULL THEN LENGTH(EXE_BLOB) ELSE 0 END FROM SISTEMA_CONFIG WHERE ROWNUM=1")
row = cur.fetchone()
print(f'Banco: VERSAO_EXIGIDA={row[0]}  EXE_BLOB={row[1]/1_048_576:.1f} MB')

cur.close()
conn.close()
print('v3.0 publicado no banco.')
