# PC Maintenance Monthly Task Scheduler

Aplikasi untuk menjalankan tugas pemeliharaan PC secara otomatis setiap bulan menggunakan Windows Task Scheduler.

## Fitur

- ✅ Berjalan otomatis pada minggu pertama bulan (tanggal 1-7)
- ✅ Hanya berjalan sekali per bulan
- ✅ Memeriksa koneksi internet sebelum menjalankan
- ✅ Kompatibel dengan Windows 7 - Windows 11
- ✅ PowerShell 2.0+ support
- ✅ Automatic retry jika gagal
- ✅ Logging dan marker tracking

## Instalasi

### Persyaratan
- Windows 7 atau lebih baru
- PowerShell 2.0 atau lebih baru
- Administrator privileges
- Koneksi Internet

### Langkah Instalasi

1. **Buka PowerShell sebagai Administrator**
   - Klik kanan pada PowerShell
   - Pilih "Run as Administrator"

2. **Navigasi ke folder aplikasi:**
   ```powershell
   cd C:\sys
   ```

3. **Jalankan script instalasi:**
   ```powershell
   .\install-monthly-task.ps1
   ```

4. **Ikuti petunjuk di layar**
   - Script akan membuat wrapper PowerShell
   - Task Scheduler akan dikonfigurasi secara otomatis
   - Anda dapat menguji task sekarang atau nanti

## Penggunaan

### Menjalankan Task Secara Manual
```powershell
schtasks /Run /TN "PC Maintenance Monthly Report"
```

### Melihat Status Task
```powershell
schtasks /Query /TN "PC Maintenance Monthly Report" /FO LIST /V
```

### Menghapus Task
```powershell
schtasks /Delete /TN "PC Maintenance Monthly Report" /F
```

## Konfigurasi

Ubah waktu eksekusi dengan parameter:
```powershell
.\install-monthly-task.ps1 -Time "14:00"
```

Parameter yang tersedia:
- `-ExePath`: Path ke MaintenanceApp.exe (default: `.\bin\Release\MaintenanceApp.exe`)
- `-TaskName`: Nama task di Task Scheduler (default: `PC Maintenance Monthly Report`)
- `-Time`: Waktu eksekusi HH:mm (default: `12:00`)
