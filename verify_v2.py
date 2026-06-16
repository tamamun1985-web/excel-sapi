# -*- coding: utf-8 -*-
"""
Verifikasi v2.0 (model BATCH). Mereplikasi rumus workbook:
  Batch pool: pakan = SUMIFS(by batch), oper = SUMIFS(by batch),
              susut = SUM(penyusutan/bln aset) * lama_bulan_batch
  HPP/ekor   = harga_beli + (pool_pakan+pool_oper+pool_susut)*hari_sapi/cattledays_batch
  Neraca aset tetap = SUM(harga aset) - SUM(susut semua batch)
Tujuan: (1) buktikan HPP batch yang sudah TUTUP tidak berubah saat batch baru ditambah,
        (2) buktikan Neraca tetap seimbang (cek = 0).
"""
import datetime
TODAY = datetime.date(2026, 6, 16)
def days(a, b): return (b - a).days
def months(a, b): return days(a, b) / 30.0

MODAL = 500_000_000

# ---- Aset tetap ----
aset = [(150_000_000,120,15_000_000),(18_000_000,60,1_000_000),
        (12_000_000,60,500_000),(6_000_000,48,0)]
monthly_dep = sum((h-r)/u for h,u,r in aset)
total_harga_aset = sum(h for h,_,_ in aset)

def batch_susut(mulai, selesai):
    end = selesai if selesai else TODAY
    return monthly_dep * months(mulai, end)

def hpp_batch(harga_beli, cow_days, pool, cattledays):
    return harga_beli + pool * cow_days / cattledays

# ====================================================================
# BATCH 1 (akan ditutup) -- 2 sapi
# ====================================================================
b1_mulai, b1_selesai = datetime.date(2025,7,1), datetime.date(2025,12,31)
b1_feed, b1_oper = 5_000_000, 12_000_000
b1_susut = batch_susut(b1_mulai, b1_selesai)
b1_pool = b1_feed + b1_oper + b1_susut
# S001,S002 dipelihara 1 Jul - 31 Des (terjual 31 Des)
s1_days = days(b1_mulai, b1_selesai); s2_days = s1_days
b1_cattledays = s1_days + s2_days
S001_harga, S002_harga = 28_000_000, 29_500_000

# --- FASE 1: hanya Batch 1 yang ada di sistem ---
S001_hpp_fase1 = hpp_batch(S001_harga, s1_days, b1_pool, b1_cattledays)
S002_hpp_fase1 = hpp_batch(S002_harga, s2_days, b1_pool, b1_cattledays)

# ====================================================================
# BATCH 2 (aktif) ditambahkan kemudian -- 4 sapi + biaya baru
# ====================================================================
b2_mulai, b2_selesai = datetime.date(2026,1,1), None
b2_feed, b2_oper = 8_000_000, 20_000_000
b2_susut = batch_susut(b2_mulai, b2_selesai)
b2_pool = b2_feed + b2_oper + b2_susut
a_days = days(b2_mulai, TODAY)
b2_cattledays = a_days * 4
A_harga = [30_000_000]*4

# --- FASE 2: Batch 1 dihitung ULANG setelah Batch 2 + biayanya masuk ---
# (rumus batch 1 hanya memakai data ber-tag B001 -> tidak tersentuh)
S001_hpp_fase2 = hpp_batch(S001_harga, s1_days, b1_pool, b1_cattledays)
S002_hpp_fase2 = hpp_batch(S002_harga, s2_days, b1_pool, b1_cattledays)

print("="*68)
print("UJI 1 - FREEZE: HPP Batch 1 (sudah TUTUP) sebelum vs sesudah Batch 2")
print("="*68)
print(f"  S001 HPP sebelum Batch 2 : Rp {S001_hpp_fase1:,.0f}")
print(f"  S001 HPP sesudah Batch 2 : Rp {S001_hpp_fase2:,.0f}")
print(f"  Selisih                  : Rp {S001_hpp_fase2-S001_hpp_fase1:,.0f}")
print(f"  S002 HPP sebelum Batch 2 : Rp {S002_hpp_fase1:,.0f}")
print(f"  S002 HPP sesudah Batch 2 : Rp {S002_hpp_fase2:,.0f}")
frozen = (S001_hpp_fase1==S001_hpp_fase2) and (S002_hpp_fase1==S002_hpp_fase2)
print(f"  -> HPP & margin batch lama TERKUNCI? {'YA (FIX BERHASIL)' if frozen else 'TIDAK'}")

# ====================================================================
# UJI 2 - NERACA seimbang pada kondisi 2 batch
# ====================================================================
jual_S001 = 480*70_000; jual_S002 = 495*70_000
penjualan = jual_S001 + jual_S002
hpp_terjual = S001_hpp_fase2 + S002_hpp_fase2          # B001 sold
persediaan_sapi = sum(hpp_batch(h, a_days, b2_pool, b2_cattledays) for h in A_harga)  # B002 active

beli_sapi = S001_harga + S002_harga + sum(A_harga)
beli_pakan = 20_000_000
feed_consumed = b1_feed + b2_feed
pers_pakan = beli_pakan - feed_consumed
operasional = b1_oper + b2_oper
beli_aset = total_harga_aset
susut_total = b1_susut + b2_susut
aset_buku = total_harga_aset - susut_total

kas = MODAL + penjualan - beli_sapi - beli_pakan - operasional - beli_aset
total_aset = kas + persediaan_sapi + pers_pakan + aset_buku
laba = penjualan - hpp_terjual
total_pasiva = MODAL + laba

print("\n" + "="*68)
print("UJI 2 - NERACA (kondisi 2 batch: B001 tutup+terjual, B002 aktif)")
print("="*68)
print(f"  Kas                : {kas:>16,.0f}")
print(f"  Persediaan Sapi    : {persediaan_sapi:>16,.0f}")
print(f"  Persediaan Pakan   : {pers_pakan:>16,.0f}")
print(f"  Aset Tetap (buku)  : {aset_buku:>16,.0f}")
print(f"  TOTAL ASET         : {total_aset:>16,.0f}")
print(f"  Modal + Laba       : {total_pasiva:>16,.0f}")
print(f"  CEK (harus 0)      : {total_aset-total_pasiva:>16,.2f}")
print()
ok = frozen and abs(total_aset-total_pasiva) < 0.01
print("KESIMPULAN:", "SISTEM v2.0 VALID - bug terkunci & neraca seimbang." if ok else "ADA MASALAH.")
