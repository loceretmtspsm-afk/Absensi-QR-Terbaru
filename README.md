ABSENSI TERPADU MTs PSM LOCERET — PAKET SCANNER GITHUB BARU

Isi paket:
- index.html : frontend scanner untuk repositori GitHub BARU. Jangan unggah ke repositori aplikasi lama.
- Code.gs    : backend proyek baru, dengan doPost untuk menerima scan dari GitHub. Modul database lain yang ada pada backend dasar tetap dipertahankan.

SEBELUM MENGGUNAKAN:
1. Pastikan Code.gs ini dipasang pada proyek Apps Script yang terhubung ke database baru SISWA, GURU, ABSENSI_SISWA, ABSENSI_GURU dan PENGATURAN.
2. Pada sheet PENGATURAN, tambahkan satu baris baru:
   KUNCI = KODE_SCANNER
   NILAI = buat kode akses acak yang panjang (disarankan 12 karakter atau lebih)
   KETERANGAN = Kode akses scanner petugas
   Jangan menaruh kode akses tersebut di file GitHub.
3. Simpan Code.gs. Di Apps Script, deploy sebagai Web App. Execute as: Me. Untuk scanner GitHub publik mengirim POST, pilih akses sesuai kebijakan sekolah; endpoint menulis absensi, jadi jangan bagikan URL dan kode akses sembarangan.
4. Salin URL web app yang berakhiran /exec.
5. Buat repositori GitHub BARU. Upload index.html ke root repositori, aktifkan GitHub Pages pada branch/folder yang dipakai, lalu buka URL GitHub Pages melalui HTTPS.
6. Di halaman scanner, tempel URL /exec backend baru dan masukkan kode akses yang cocok dengan KODE_SCANNER di PENGATURAN.
7. Setelah mengubah Code.gs, deploy versi baru (Manage deployments > Edit > New version > Deploy).

CATATAN PENTING:
- Jangan mengubah atau menimpa repositori scanner lama yang masih digunakan.
- QR siswa lama NIS:nomorinduk didukung. QR guru dicocokkan dengan KODE_QR yang tersimpan pada sheet GURU.
- QR siswa memerlukan izin lokasi; siswa di luar radius sekolah akan ditolak.
- Waktu siswa: MASUK mulai 07.00 WIB, PULANG mulai 13.00 WIB. Jika siswa tidak memiliki catatan MASUK hari itu, backend menolak mencatat PULANG langsung.
- Pemeriksaan yang dilakukan di lingkungan ini hanya pemeriksaan sintaks JavaScript dan struktur kode. Integrasi runtime dengan Google Apps Script dan spreadsheet harus dikonfirmasi setelah dipasang; jangan menganggap scan berhasil sampai hasilnya terlihat di aplikasi dan baris absensi yang benar muncul di spreadsheet.
