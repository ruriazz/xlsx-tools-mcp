MCP server lokal berbasis python
kebutuhan untuk baca dan edit file excel secara akurat

Stack final buat MCP server read+write xlsx (Python):

**Core:**
1. **`openpyxl`** — baca & tulis utama. Preserve formula, style, merged cells, comment. Ini satu-satunya library yang bisa edit file existing tanpa ngerusak struktur aslinya.
2. **`pandas`** — transform/agregasi data bulk (dipakai di atas hasil baca openpyxl, bukan gantiin).
3. **LibreOffice headless (`soffice --headless --convert-to xlsx --calc`)** — wajib buat recalculate formula setelah openpyxl nulis, karena openpyxl nggak eval formula sama sekali (nulis string doang, cached value kosong). Tanpa ini, hasil tulis kamu keliatan "kosong" di setiap tool yang baca cached value (pandas, Excel viewer, dsb).
4. **`python-calamine`** — dipakai sebagai read-path untuk file besar/kompleks atau sebagai cross-check saat parsing mencurigakan (mismatched type, dsb). Read-only, jauh lebih cepat & robust dari openpyxl di file besar.

**Support (non-negotiable buat production-grade):**
5. **`defusedxml`** — xlsx = zip+XML, wajib biar nggak vulnerable ke XML entity expansion attack dari file upload user.
6. **`filelock`** — kalau MCP server bakal handle concurrent read/write ke file yang sama, wajib biar nggak corrupt.

Itu saja — enam library, tidak ada cabang "kalau butuh X pakai Y". Alur kerjanya baku:
- **Write**: openpyxl tulis → LibreOffice recalc → validasi `errors_found` dari output JSON recalc sebelum return sukses ke caller.

> **Catatan (tradeoff recalc):** round-trip LibreOffice headless itu *tradeoff optimasi*, bukan jaminan
> bit-for-bit lossless. LibreOffice menghitung ulang formula tapi mengekspor ulang *seluruh workbook*,
> sehingga fitur yang dipertahankan openpyxl belum tentu dipertahankan identik (pivot table, chart,
> data validation, sebagian format / defined names). Di workbook yang strukturnya kompleks dan
> resiko ini penting, lewati round-trip dengan `recalculate=False` pada tool write
> (`write_cells`, `append_rows`, `insert/delete_rows`, `insert/delete_columns`) supaya ditulis hanya
> dengan openpyxl tanpa recalc.
- **Read**: calamine sebagai primary parser (cepat, akurat tipe data) → fallback openpyxl kalau perlu formula string/style metadata yang calamine gak expose.
- **Bulk transform**: pandas di atas hasil baca, bukan baca langsung dari file (biar merged cells & style nggak keburu ke-flatten salah).

---

## Struktur project

```
xlsx-reader/
├── pyproject.toml
├── readme.md
└── src/xlsx_mcp/
    ├── server.py          # entry point, definisi 19 MCP tools
    ├── recalc.py           # LibreOffice headless recalc + scan error formula
    ├── locking.py           # file lock per-path (filelock)
    ├── settings.py          # timeout recalc & lock (env var)
    ├── errors.py            # exception domain (SheetNotFoundError, dst.)
    └── io/
        ├── reader.py         # calamine primary + openpyxl fallback
        ├── writer.py         # openpyxl write pipeline + trigger recalc
        └── transform.py      # agregasi pandas di atas hasil reader
```

Read dan write dipisah total dari transport MCP (`server.py` cuma wrapper tipis + validasi + file lock) supaya logic-nya bisa dites/dipakai ulang tanpa jalanin server.

## Instalasi

Butuh **Python ≥ 3.10** dan **LibreOffice** (untuk recalc formula — tanpa ini, write tetap jalan tapi formula gak kehitung ulang, cuma dikasih warning di response).

```bash
# LibreOffice (wajib untuk recalc akurat)
brew install --cask libreoffice        # macOS
sudo apt-get install -y libreoffice-calc   # Debian/Ubuntu

# Install project + dependency
cd xlsx-reader
uv sync            # atau: pip install -e .
```

Cek server jalan:

```bash
uv run xlsx-mcp
# atau
uv run python -m xlsx_mcp
```

Server jalan pakai stdio transport (standar MCP) — proses ini nunggu client (Claude Code/OpenCode) yang connect, bukan buat dijalanin manual & dibiarin nyala di terminal.

## Konfigurasi di Claude Code

```bash
claude mcp add xlsx-mcp -- uv --directory /path/absolut/ke/xlsx-reader run xlsx-mcp
```

Atau tambah manual di `.claude.json` / `.mcp.json` project:

```json
{
  "mcpServers": {
    "xlsx-mcp": {
      "command": "uv",
      "args": ["--directory", "/path/absolut/ke/xlsx-reader", "run", "xlsx-mcp"]
    }
  }
}
```

## Konfigurasi di OpenCode

Tambah di `opencode.json` (project) atau `~/.config/opencode/opencode.json` (global):

```json
{
  "mcp": {
    "xlsx-mcp": {
      "type": "local",
      "command": ["uv", "--directory", "/path/absolut/ke/xlsx-reader", "run", "xlsx-mcp"],
      "enabled": true
    }
  }
}
```

Ganti `/path/absolut/ke/xlsx-reader` sesuai lokasi clone-nya. Kalau tidak pakai `uv`, bisa juga `"command": ["python", "-m", "xlsx_mcp"]` asal virtualenv-nya sudah aktif/di-resolve dengan benar.

## Preload file (biar agent gak perlu cari path)

Set env var `XLSX_MCP_FILES` di config MCP server (bukan di shell) supaya file-file yang bakal dipakai sudah "dikenal" server sejak start — semua tool tinggal dipanggil pakai nama itu (atau tanpa `path` sama sekali kalau cuma 1 file) tanpa agent perlu nyari-nyari di disk duluan.

Format: comma-separated, tiap entri `nama=/path/absolut.xlsx` (atau bare path — nama otomatis diambil dari nama file):

```json
{
  "mcpServers": {
    "xlsx-mcp": {
      "command": "uv",
      "args": ["--directory", "/path/absolut/ke/xlsx-reader", "run", "xlsx-mcp"],
      "env": {
        "XLSX_MCP_FILES": "e2e.xlsx=/data/e2e.xlsx,laporan=/data/laporan-bulanan.xlsx"
      }
    }
  }
}
```

- Cuma 1 file dikonfigurasi → semua tool bisa dipanggil tanpa argumen `path` sama sekali.
- Lebih dari 1 file → panggil tool dengan `path` = nama alias-nya (mis. `"path": "e2e.xlsx"`), bukan path fisiknya.
- Tool `list_configured_files()` mengembalikan mapping alias → path fisik yang ter-preload, buat agent cek dulu kalau ragu.
- `path` tetap bisa diisi path fisik biasa (absolut/relatif) untuk file di luar daftar preload — itu tetap jalan seperti biasa.

## Daftar tools

Semua path bisa absolut atau relatif ke working directory server. Semua tool **write** mengembalikan `{saved, recalculated, errors_found, message}` — selalu cek `errors_found` sebelum anggap tulisan berhasil bersih (bisa saja tersimpan tapi formula-nya menghasilkan `#DIV/0!` dsb).

Semua `path` **opsional** kalau file sudah di-preload lewat `XLSX_MCP_FILES` (lihat bagian di atas) — omit kalau cuma 1 file configured, atau isi nama alias-nya kalau lebih dari 1.

**Baca / inspeksi**
| Tool | Kegunaan |
|---|---|
| `list_configured_files()` | Daftar file yang ter-preload saat startup (alias → path fisik) |
| `list_sheets(path?)` | Daftar sheet + estimasi ukuran (cepat, calamine) |
| `get_workbook_info(path)` | Metadata workbook: dimensi presisi, active sheet, defined names |
| `read_sheet(path, sheet, cell_range?, max_rows?)` | Baca nilai sel sebagai array 2D, koordinat absolut dari A1 |
| `get_cell(path, sheet, cell)` | Detail 1 sel: value, formula, number_format, font, fill, merge, comment |
| `search_workbook(path, query, sheet?, match_case?, limit?)` | Cari substring di seluruh/1 sheet |
| `aggregate_sheet(path, sheet, group_by, agg, cell_range?, has_header?)` | Group + agregasi pakai pandas (mis. sum per kategori) |

**Tulis**
| Tool | Kegunaan |
|---|---|
| `create_workbook(path, sheets?, overwrite?)` | Buat file .xlsx baru — `path` wajib diakhiri `.xlsx`/`.xlsm` |
| `write_cells(path, sheet, cells, create_sheet_if_missing?, recalculate?)` | Tulis value/formula ke sel spesifik, lalu recalc |
| `append_rows(path, sheet, rows, create_sheet_if_missing?, recalculate?)` | Tambah baris di akhir data, lalu recalc |
| `create_sheet(path, sheet, index?)` | Tambah sheet kosong |
| `delete_sheet(path, sheet)` | Hapus sheet (gagal kalau tinggal 1 sheet) |
| `insert_rows(path, sheet, start_row, count?, recalculate?)` | Sisip baris kosong |
| `delete_rows(path, sheet, start_row, count?, recalculate?)` | Hapus baris |
| `insert_columns(path, sheet, start_column, count?, recalculate?)` | Sisip kolom kosong |
| `delete_columns(path, sheet, start_column, count?, recalculate?)` | Hapus kolom |
| `merge_cells(path, sheet, cell_range)` / `unmerge_cells(...)` | Gabung/pisah sel |
| `set_cell_style(path, sheet, cell_range, style)` | Format sel: `bold, italic, font_size, font_color, bg_color, horizontal, vertical, border, number_format` |
| `recalculate_workbook(path)` | Paksa recalc LibreOffice + laporkan error formula tanpa nulis apa-apa |

### Contoh pemanggilan tool (payload JSON)

```json
{
  "tool": "write_cells",
  "arguments": {
    "path": "/data/laporan.xlsx",
    "sheet": "Ringkasan",
    "cells": [
      { "cell": "A1", "value": "Total" },
      { "cell": "B1", "formula": "=SUM(Data!B2:B100)" }
    ]
  }
}
```

Response-nya:

```json
{
  "saved": true,
  "recalculated": true,
  "errors_found": [],
  "message": "Recalculated with LibreOffice headless."
}
```

Kalau `errors_found` tidak kosong, isinya `[{"sheet": "...", "cell": "...", "error": "#DIV/0!"}]` — artinya file tetap tersimpan, tapi ada formula yang gagal dihitung dan perlu ditinjau.

## Keamanan & konkurensi

- **XML bomb protection**: otomatis aktif — `openpyxl` mendeteksi package `defusedxml` terpasang dan pakai parser-nya, tanpa konfigurasi tambahan.
- **File lock**: setiap tool baca/tulis mengunci `<path>.lock` (via `filelock`) selama operasi berlangsung, supaya panggilan tool yang tumpang tindih pada file yang sama tidak saling korup.
- Env var opsional: `XLSX_MCP_RECALC_TIMEOUT` (default 60 detik), `XLSX_MCP_LOCK_TIMEOUT` (default 10 detik).

## Troubleshooting

- **`errors_found` selalu kosong padahal ada formula rusak** → kemungkinan LibreOffice belum ter-install; cek pesan di field `message` pada response write, biasanya eksplisit bilang "LibreOffice (soffice) not found on PATH".
- **Recalc lambat / timeout di file besar** → naikkan `XLSX_MCP_RECALC_TIMEOUT`.
- **`LockTimeoutError`** → ada operasi lain yang masih memegang lock file yang sama; tunggu atau naikkan `XLSX_MCP_LOCK_TIMEOUT`.
- **Sheet tidak ketemu** → semua tool melempar pesan error yang eksplisit menyebutkan nama sheet yang tersedia di file tersebut.
