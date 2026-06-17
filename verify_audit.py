# -*- coding: utf-8 -*-
"""
Audit klaim: 'biaya operasional periode baru mengubah HPP & margin sapi yang sudah terjual'.
Mereplikasi rumus AKTUAL workbook:
  HPP_Sapi!Oper  = SUM(Operasional!C SELURUH waktu) * HariSapi / SUM(M_Sapi!Hari SELURUH waktu)
  Penjualan!HPP  = INDEX(HPP_Sapi!HPP ; live lookup)  -> tidak terkunci
"""

def oper_alloc(total_oper_all, cow_days, total_days_all):
    return total_oper_all * cow_days / total_days_all

print("="*70)
print("SKENARIO: Batch 1 (S001-S004) dijual 31 Mar, lalu Batch 2 (A001-A004) Mei")
print("="*70)

# ---- Batch 1 ----
b1_days_each = 30          # 1-31 Mar
b1_n = 4
b1_total_days = b1_days_each * b1_n         # 120
oper_maret = 2_060_000
harga_beli_S001 = 28_000_000

# Saat hanya Batch 1 ada (evaluasi akhir Maret, sebelum Batch 2 diinput)
total_oper_A = oper_maret
total_days_A = b1_total_days
S001_oper_A  = oper_alloc(total_oper_A, b1_days_each, total_days_A)
S001_hpp_A   = harga_beli_S001 + S001_oper_A     # (abaikan pakan/penyusutan utk fokus)

print("\n--- KONDISI 1: hanya Batch 1 (saat S001 DIJUAL 31 Mar) ---")
print(f"  Total operasional tercatat : Rp {total_oper_A:,.0f}")
print(f"  Total hari semua sapi      : {total_days_A} hari")
print(f"  Alokasi operasional S001   : Rp {S001_oper_A:,.0f}")
print(f"  HPP S001 (saat dijual)     : Rp {S001_hpp_A:,.0f}")

# ---- Tambah Batch 2 (Mei) ----
b2_days_each = 30          # 1-31 Mei
b2_n = 4
b2_total_days = b2_days_each * b2_n         # 120
oper_mei = 1_060_000

total_oper_B = oper_maret + oper_mei
total_days_B = b1_total_days + b2_total_days
S001_oper_B  = oper_alloc(total_oper_B, b1_days_each, total_days_B)   # S001 hari TETAP 30
S001_hpp_B   = harga_beli_S001 + S001_oper_B

print("\n--- KONDISI 2: setelah Batch 2 (A001-A004) + operasional Mei diinput ---")
print(f"  Total operasional tercatat : Rp {total_oper_B:,.0f}  (Maret + Mei)")
print(f"  Total hari semua sapi      : {total_days_B} hari")
print(f"  Alokasi operasional S001   : Rp {S001_oper_B:,.0f}   <-- BERUBAH")
print(f"  HPP S001 (padahal sudah laku 2 bln lalu) : Rp {S001_hpp_B:,.0f}")

selisih = S001_hpp_B - S001_hpp_A
print("\n" + "="*70)
print("HASIL AUDIT")
print("="*70)
print(f"  HPP S001 SEBELUM input Batch 2 : Rp {S001_hpp_A:,.0f}")
print(f"  HPP S001 SESUDAH input Batch 2 : Rp {S001_hpp_B:,.0f}")
print(f"  PERUBAHAN HPP sapi yg sdh laku : Rp {selisih:,.0f}")
print(f"  -> Margin penjualan Maret ikut bergeser sebesar Rp {abs(selisih):,.0f}")
print()
if abs(selisih) > 0:
    print("  KESIMPULAN: TERBUKTI. HPP & margin sapi yang sudah terjual TIDAK")
    print("  terkunci dan berubah saat ada biaya/sapi periode baru. Klaim BENAR.")
else:
    print("  Tidak ada perubahan.")
