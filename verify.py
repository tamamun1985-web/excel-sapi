# -*- coding: utf-8 -*-
"""Verifikasi independen logika perhitungan (replikasi formula Excel pada data contoh)."""
import datetime

TODAY = datetime.date(2026, 6, 11)  # sesuai konteks tanggal sistem

def datedif_m(a, b):  # complete months, seperti Excel DATEDIF(...,"m")
    m = (b.year - a.year) * 12 + (b.month - a.month)
    if b.day < a.day:
        m -= 1
    return m

def days(a):
    return (TODAY - a).days

MODAL = 500_000_000

# --- Aset tetap ---
aset = [
    ("Kandang",   datetime.date(2024,1,10), 150_000_000, 120, 15_000_000),
    ("Chopper",   datetime.date(2024,1,15),  18_000_000,  60,  1_000_000),
    ("Timbangan", datetime.date(2024,2,1),   12_000_000,  60,    500_000),
    ("Karpet",    datetime.date(2024,2,1),    6_000_000,  48,          0),
]
total_harga_aset = sum(a[2] for a in aset)
akm = 0.0
for nm,tgl,harga,umur,residu in aset:
    pb = (harga-residu)/umur
    mused = datedif_m(tgl, TODAY)
    a_akm = min(harga-residu, pb*mused)
    akm += a_akm
nilai_buku_aset = total_harga_aset - akm

# --- Sapi ---
sapi = [  # id, masuk, harga, kelompok, bobot_awal, bobot_terkini
    ("S001", datetime.date(2025,9,1), 28_000_000, "K1", 280, 312),
    ("S002", datetime.date(2025,9,1), 29_500_000, "K1", 290, 325),
    ("S003", datetime.date(2025,8,15),31_000_000, "K2", 330, 372),
    ("S004", datetime.date(2025,8,15),33_000_000, "K2", 345, 388),
    ("S005", datetime.date(2025,7,1), 41_000_000, "K3", 410, 458),
    ("S006", datetime.date(2025,7,1), 42_500_000, "K3", 420, 470),
]
beli_sapi = sum(s[2] for s in sapi)
hari = {s[0]: days(s[1]) for s in sapi}
total_cattledays = sum(hari.values())

# --- Pakan masuk (rata-rata tertimbang per kode) ---
pakan_masuk = [("PK01",3000,1_500_000),("PK02",1000,3_500_000),("PK03",50,750_000),
               ("PK01",3000,1_560_000),("PK02",1200,4_320_000)]
masuk_qty, masuk_val = {}, {}
for kode,q,v in pakan_masuk:
    masuk_qty[kode]=masuk_qty.get(kode,0)+q
    masuk_val[kode]=masuk_val.get(kode,0)+v
avg = {k: masuk_val[k]/masuk_qty[k] for k in masuk_qty}
beli_pakan = sum(masuk_val.values())

# --- Pakan keluar ---
pakan_keluar = [("K1","PK01",40),("K1","PK02",12),("K2","PK01",45),
                ("K2","PK02",14),("K3","PK01",55),("K3","PK02",18)]
feed_group = {}
feed_total = 0.0
for kel,kode,q in pakan_keluar:
    nilai = q*avg[kode]
    feed_group[kel]=feed_group.get(kel,0)+nilai
    feed_total += nilai
keluar_qty = {}
for kel,kode,q in pakan_keluar:
    keluar_qty[kode]=keluar_qty.get(kode,0)+q

# --- Operasional ---
oper_total = 4_500_000+650_000+500_000+800_000

# --- HPP per ekor ---
group_days = {}
for s in sapi:
    group_days[s[3]] = group_days.get(s[3],0)+hari[s[0]]
print("=== HPP PER EKOR ===")
hpp = {}
persediaan_sapi = 0.0
for sid,masuk,harga,kel,b0,bt in sapi:
    d = hari[sid]
    feed_alloc = feed_group.get(kel,0)*d/group_days[kel]
    oper_alloc = oper_total*d/total_cattledays
    susut_alloc = akm*d/total_cattledays
    h = harga+feed_alloc+oper_alloc+susut_alloc
    hpp[sid]=h
    persediaan_sapi += h  # semua aktif
    print(f"{sid} {kel} hari={d:3d}  beli={harga:>12,.0f}  pakan={feed_alloc:>10,.0f}  "
          f"oper={oper_alloc:>10,.0f}  susut={susut_alloc:>11,.0f}  HPP={h:>13,.0f}")

# --- Stok pakan ---
pers_pakan = sum((masuk_qty[k]-keluar_qty.get(k,0))*avg[k] for k in masuk_qty)

# --- Laporan ---
pendapatan = 0.0; hpp_terjual=0.0; laba=pendapatan-hpp_terjual
kas = MODAL + pendapatan - beli_sapi - beli_pakan - oper_total - total_harga_aset
total_aset = kas + persediaan_sapi + pers_pakan + nilai_buku_aset
total_pasiva = MODAL + laba

print("\n=== FCG PER KELOMPOK ===")
for kel in ["K1","K2","K3"]:
    gains = sum((bt-b0) for sid,_,_,k,b0,bt in sapi if k==kel)
    fcg = feed_group[kel]/gains
    print(f"{kel}: gain={gains} kg  biaya_pakan={feed_group[kel]:,.0f}  FCG={fcg:,.0f} Rp/kg")

print("\n=== ADG PER EKOR ===")
for sid,masuk,harga,kel,b0,bt in sapi:
    print(f"{sid}: gain={bt-b0} kg / {hari[sid]} hari = ADG {(bt-b0)/hari[sid]:.3f} kg/hari")

print("\n=== NERACA / KESEIMBANGAN ===")
print(f"Kas               : {kas:>16,.0f}")
print(f"Persediaan Sapi   : {persediaan_sapi:>16,.0f}")
print(f"Persediaan Pakan  : {pers_pakan:>16,.0f}")
print(f"Aset Tetap (buku) : {nilai_buku_aset:>16,.0f}")
print(f"TOTAL ASET        : {total_aset:>16,.0f}")
print(f"Modal+Laba        : {total_pasiva:>16,.0f}")
print(f"CEK (harus 0)     : {total_aset-total_pasiva:>16,.2f}")
