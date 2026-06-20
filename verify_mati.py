# -*- coding: utf-8 -*-
"""
Audit modul SAPI MATI/AFKIR. Mereplikasi rumus workbook pada data contoh
(1 batch B001: S001-S005 aktif, S006 MATI 15/01/2026).
Membuktikan: kerugian = HPP sapi mati, masuk Laba Rugi, keluar dari Persediaan,
tidak ada double counting, dan Neraca tetap balance = 0.
"""
import datetime
TODAY = datetime.date(2026, 6, 19)
def days(a, b): return (b - a).days

MODAL = 500_000_000

# ---- Aset (penyusutan/bln) ----
aset = [(150_000_000,120,15_000_000),(18_000_000,60,1_000_000),
        (12_000_000,60,500_000),(6_000_000,48,0)]
monthly_dep = sum((h-r)/u for h,u,r in aset)
total_harga_aset = sum(h for h,_,_ in aset)

# ---- Batch B001: mulai 1/7/2025, aktif ----
b_mulai = datetime.date(2025,7,1)
b_bulan = days(b_mulai, TODAY)/30.0
b_susut = monthly_dep * b_bulan

# ---- Pakan masuk (rata-rata) ----
avg = {'PK01':3_060_000/6000, 'PK02':7_820_000/2200, 'PK03':15000}
# Pakan keluar B001
feed = 140*avg['PK01'] + 44*avg['PK02']          # PK01:40+45+55, PK02:12+14+18
oper = 4_500_000+650_000+500_000+800_000
pool = feed + oper + b_susut

# ---- Sapi ----
sapi = [  # id, masuk, harga, status, tgl_keluar(mati)
 ('S001',datetime.date(2025,9,1),28_000_000,'Aktif',None),
 ('S002',datetime.date(2025,9,1),29_500_000,'Aktif',None),
 ('S003',datetime.date(2025,8,15),31_000_000,'Aktif',None),
 ('S004',datetime.date(2025,8,15),33_000_000,'Aktif',None),
 ('S005',datetime.date(2025,7,1),41_000_000,'Aktif',None),
 ('S006',datetime.date(2025,7,1),42_500_000,'Mati',datetime.date(2026,1,15)),
]
def hari(s):
    end = s[4] if s[4] else TODAY
    return days(s[1], end)
cattledays = sum(hari(s) for s in sapi)

print("="*70)
print("AUDIT SAPI MATI -- HPP per ekor (S006 = MATI 15/01/2026)")
print("="*70)
hpp = {}
for s in sapi:
    d = hari(s)
    pakan_alok = feed*d/cattledays
    oper_alok  = oper*d/cattledays
    susut_alok = b_susut*d/cattledays
    h = s[2] + pakan_alok + oper_alok + susut_alok
    hpp[s[0]] = h
    tag = "  <== MATI" if s[3]=="Mati" else ""
    print(f"  {s[0]} ({s[3]:6}) hari={d:3d}  HPP=Rp {h:>13,.0f}{tag}")
    if s[3]=="Mati":
        print(f"        rincian: beli {s[2]:,.0f} + pakan {pakan_alok:,.0f} + oper {oper_alok:,.0f} + susut {susut_alok:,.0f}")

kerugian = sum(hpp[s[0]] for s in sapi if s[3]=='Mati')
persediaan_sapi = sum(hpp[s[0]] for s in sapi if s[3]=='Aktif')   # MATI tidak masuk
penjualan = 0
hpp_terjual = 0
laba = penjualan - hpp_terjual - kerugian

# ---- Neraca ----
beli_sapi = sum(s[2] for s in sapi)
beli_pakan = 11_630_000
pers_pakan = beli_pakan - feed
operasional = oper
beli_aset = total_harga_aset
aset_buku = total_harga_aset - b_susut
kas = MODAL + penjualan - beli_sapi - beli_pakan - operasional - beli_aset
total_aset = kas + persediaan_sapi + pers_pakan + aset_buku
total_pasiva = MODAL + laba

print("\n" + "="*70)
print("DAMPAK KE LAPORAN")
print("="*70)
print(f"  [Laba Rugi]  Kerugian Sapi Mati (HPP S006) : Rp {kerugian:,.0f}")
print(f"  [Laba Rugi]  Laba Bersih (0 - kerugian)     : Rp {laba:,.0f}")
print(f"  [Neraca]     Persediaan Sapi (5 aktif saja) : Rp {persediaan_sapi:,.0f}")
print(f"  [Neraca]     S006 dihitung sbg persediaan?  : TIDAK (no double counting)")
print(f"  [Dashboard]  Sapi Mati/Afkir                : {sum(1 for s in sapi if s[3]=='Mati')} ekor")
print(f"  [Dashboard]  Laba Bersih mencerminkan rugi  : Rp {laba:,.0f}")
print("\n  --- NERACA ---")
print(f"  Kas               : {kas:>16,.0f}")
print(f"  Persediaan Sapi   : {persediaan_sapi:>16,.0f}")
print(f"  Persediaan Pakan  : {pers_pakan:>16,.0f}")
print(f"  Aset Tetap (buku) : {aset_buku:>16,.0f}")
print(f"  TOTAL ASET        : {total_aset:>16,.0f}")
print(f"  Modal + Laba      : {total_pasiva:>16,.0f}")
print(f"  CEK (harus 0)     : {total_aset-total_pasiva:>16,.2f}")
ok = abs(total_aset-total_pasiva)<0.01 and kerugian>0
print("\nKESIMPULAN:", "Modul sapi mati BENAR - rugi tercatat, no double count, neraca seimbang." if ok else "MASALAH.")
