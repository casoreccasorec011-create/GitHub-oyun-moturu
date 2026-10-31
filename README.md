# GitHub-oyun-moturu
GitHub oyun moturu
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# Create workbook
wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Colors (Dark Slate / Tech Cyan Theme)
HEADER_BG = "1F4E78" # Dark Blue/Slate
HEADER_FG = "FFFFFF"
ROW_LIGHT = "F2F7FA"
ROW_WHITE = "FFFFFF"
ACCENT_BG = "D9E1F2"
BORDER_COLOR = "D9D9D9"
HIGHLIGHT_FILL = "E2EFDA" # Soft light green

font_title = Font(name="Segoe UI", size=16, bold=True, color="1F4E78")
font_subtitle = Font(name="Segoe UI", size=11, italic=True, color="595959")
font_header = Font(name="Segoe UI", size=11, bold=True, color=HEADER_FG)
font_bold = Font(name="Segoe UI", size=10, bold=True)
font_regular = Font(name="Segoe UI", size=10)

fill_header = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")
fill_light = PatternFill(start_color=ROW_LIGHT, end_color=ROW_LIGHT, fill_type="solid")
fill_accent = PatternFill(start_color=ACCENT_BG, end_color=ACCENT_BG, fill_type="solid")
fill_highlight = PatternFill(start_color=HIGHLIGHT_FILL, end_color=HIGHLIGHT_FILL, fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color=BORDER_COLOR),
    right=Side(style='thin', color=BORDER_COLOR),
    top=Side(style='thin', color=BORDER_COLOR),
    bottom=Side(style='thin', color=BORDER_COLOR)
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

# ---------------------------------------------------------
# TAB 1: Master Plan & Vizyon
# ---------------------------------------------------------
ws1 = wb.create_sheet(title="Vizyon & Master Plan")
ws1.views.sheetView[0].showGridLines = True

ws1["A1"] = "AI x 3D: Yeni Nesil Yapay Zeka Oyun Motoru"
ws1["A1"].font = font_title
ws1["A2"] = "Slogan: 'Fikirden nihai oyuna kadar her şey yapay zeka ile otomatik oluşturulur.'"
ws1["A2"].font = font_subtitle

headers_1 = ["Modül / Yapı", "Açıklama & Fonksiyon", "Hedef Teknoloji / Yapay Zeka", "Öncelik"]
for col_idx, header in enumerate(headers_1, start=1):
    cell = ws1.cell(row=4, column=col_idx, value=header)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border

data_1 = [
    ["Komut Girişi (Prompt to Engine)", "Kullanıcının doğal dilde yazdığı ham hikaye, mekanik ve atmosfer isteklerini kabul eden akıllı girdi paneli.", "LLM (GPT-4o / Claude / Custom Agent)", "Yüksek"],
    ["AI Core (Çekirdek)", "Metni analiz eder; sahne senaryosunu, harita büyüklüğünü, NPC davranışlarını ve görev zincirlerini saniyeler içinde planlar.", "Multi-Agent System & Graph AI", "Kritik"],
    ["Asset & Generation Pipeline", "En kaliteli 3D varlıkları, PBR dokuları, modelleri ve sesleri sahneye otomatik çeker veya procedural üretir.", "Procedural Asset Generation & Asset Store APIs", "Kritik"],
    ["Dünya ve Çevre Motoru (World Generator)", "Ormanlar, mağaralar, harabeler ve hava durumu efektlerini (Kızıl yağmur, sis, ıslanma fizikleri) otomatik oluşturur.", "Procedural Terrain & Custom Shaders", "Yüksek"],
    ["Karakter ve Hikaye Motoru", "Karakter yaratma, animasyon bağlama, Codex Rubeus gibi konuşan dinamik AI rehber ve kopan bağlantı mekanikleri.", "Interactive NPC Framework & Behavior Trees", "Yüksek"],
    ["Sığınak & Hayatta Kalma (Survival)", "Açlık/susuzluk barları, sığınak/atölye inşası, otonom savunma kuleleri ve NPC takas ekolojisi.", "Game Logic Automator & Rule Engine", "Orta"]
]

for row_idx, row_data in enumerate(data_1, start=5):
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=value)
        cell.font = font_regular
        cell.border = thin_border
        cell.alignment = align_center if col_idx in [1, 4] else align_left
        if row_idx % 2 == 0:
            cell.fill = fill_light

# ---------------------------------------------------------
# TAB 2: Teknik Mimarisi & Modüller
# ---------------------------------------------------------
ws2 = wb.create_sheet(title="Modül Detayları")
ws2.views.sheetView[0].showGridLines = True

ws2["A1"] = "Sistem Modülleri ve Teknik Gereksinimler"
ws2["A1"].font = font_title

headers_2 = ["Modül Kodu", "Modül Adı", "Girdi (Input)", "İşlem (Process)", "Çıktı (Output)", "Kullanılan AI Modeli / Kütüphane"]
for col_idx, header in enumerate(headers_2, start=1):
    cell = ws2.cell(row=3, column=col_idx, value=header)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border

data_2 = [
    ["MOD-01", "Natural Language Parser", "Kullanıcı Metin Komutu", "Prompt ayrıştırma, varlık ve mekanik listesi çıkarma", "JSON Yapılandırma Şeması", "LLM Prompt Parser"],
    ["MOD-02", "Terrain & World Gen", "Harita Büyüklüğü & Biyom Tipi", "Procedural arazi, bitki örtüsü ve hava shader'ları yükleme", "3D Sahne Arazisi & Hava Efektleri", "Procedural World Generator"],
    ["MOD-03", "Narrative & Codex AI", "Görev ve Diyalog Şablonu", "Dinamik seslendirme, laf sokan diyaloglar, Signal Loss durumu", "Diyalog Metinleri & Ses Dosyaları", "Custom NPC Agent & TTS"],
    ["MOD-04", "Asset Manager", "Model İhtiyaç Listesi", "12K PBR kaplamalar, 3D nesneler ve animasyon seçimi", "Yerleştirilmiş 3D Prefab'lar", "Asset Browser Auto-Fetcher"],
    ["MOD-05", "Survival Systems", "Açlık/Susuzluk & Sığınak Kurma", "Sığınak kurma kuralları, otonom kule AI, envanter mantığı", "Oynanabilir Oyun Mantığı (C# Scripts)", "GameLogic Code Generator"]
]

for row_idx, row_data in enumerate(data_2, start=4):
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=value)
        cell.font = font_regular
        cell.border = thin_border
        cell.alignment = align_center if col_idx in [1, 3, 5] else align_left
        if row_idx % 2 == 1:
            cell.fill = fill_light

# ---------------------------------------------------------
# TAB 3: Yol Haritası & Milestones
# ---------------------------------------------------------
ws3 = wb.create_sheet(title="Yol Haritası (Roadmap)")
ws3.views.sheetView[0].showGridLines = True

ws3["A1"] = "Geliştirme Yol Haritası ve Aşama Hedefleri"
ws3["A1"].font = font_title

headers_3 = ["Aşama (Phase)", "Aşama Adı", "Tahmini Süre", "Temel Kazanımlar / Çıktılar", "Durum"]
for col_idx, header in enumerate(headers_3, start=1):
    cell = ws3.cell(row=3, column=col_idx, value=header)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border

data_3 = [
    ["Phase 1", "Ar-Ge & Prompt Parsing Engine", "1.5 Ay", "Metinden JSON şema üreten LLM entegrasyonu ve temeller.", "Tamamlandı"],
    ["Phase 2", "3D Sahne ve Asset Generator", "2.5 Ay", "Procedural harita, hava durumu ve hazır 3D varlık çekici.", "Devam Ediyor"],
    ["Phase 3", "Codex AI & Oyun Mantığı Otomasyonu", "2 Ay", "Konuşan rehber sistemi, Signal Loss ve sığınak kurma mantığı.", "Planlandı"],
    ["Phase 4", "Entegrasyon ve Test (MVP)", "1.5 Ay", "Tüm modüllerin birleştirilmesi, testler ve UI düzenlemesi.", "Planlandı"],
    ["Phase 5", "Kapalı Beta ve Yayınlama", "1 Ay", "Kullanıcı geri bildirimleri, performans optimizasyonu.", "Planlandı"]
]

for row_idx, row_data in enumerate(data_3, start=4):
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws3.cell(row=row_idx, column=col_idx, value=value)
        cell.font = font_regular
        cell.border = thin_border
        cell.alignment = align_center if col_idx in [1, 3, 5] else align_left
        if row_idx % 2 == 1:
            cell.fill = fill_light

# ---------------------------------------------------------
# TAB 4: Bütçe ve Maliyet Analizi
# ---------------------------------------------------------
ws4 = wb.create_sheet(title="Bütçe & Tahmini Maliyet")
ws4.views.sheetView[0].showGridLines = True

ws4["A1"] = "AI Oyun Motoru Tahmini Geliştirme Bütçesi"
ws4["A1"].font = font_title

headers_4 = ["Gider Kalemi", "Kategori", "Aylık Birim Maliyet ($)", "Süre (Ay)", "Toplam Maliyet ($)"]
for col_idx, header in enumerate(headers_4, start=1):
    cell = ws4.cell(row=3, column=col_idx, value=header)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border

data_4 = [
    ["GPU Sunucu Altyapısı (AWS/RunPod)", "Sunucu & Altyapı", 1200, 8, "=C4*D4"],
    ["LLM & Voice API Kullanımı (OpenAI/ElevenLabs)", "Yapay Zeka API", 800, 8, "=C5*D5"],
    ["3D Asset & Shader Kütüphane Lisansları", "Varlık / Asset", 500, 1, "=C6*D6"],
    ["Yazılım / AI Mühendisliği (2 Kişi)", "Personel", 8000, 8, "=C7*D7"],
    ["Arayüz & UI/UX Tasarımı", "Tasarım", 2000, 3, "=C8*D8"],
    ["Test & Bulut Sunucu Dağıtımı", "Operasyon", 400, 8, "=C9*D9"]
]

for row_idx, row_data in enumerate(data_4, start=4):
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws4.cell(row=row_idx, column=col_idx, value=value)
        cell.font = font_regular
        cell.border = thin_border
        if col_idx in [3, 5]:
            cell.number_format = '$#,##0'
            cell.alignment = align_right
        elif col_idx == 4:
            cell.number_format = '#,##0'
            cell.alignment = align_center
        else:
            cell.alignment = align_left if col_idx == 1 else align_center
        
        if row_idx % 2 == 1:
            cell.fill = fill_light

# Total Row
tot_row = len(data_4) + 4
ws4.cell(row=tot_row, column=1, value="TOPLAM TAHMİNİ BÜTÇE").font = font_bold
ws4.cell(row=tot_row, column=1).alignment = align_left
ws4.cell(row=tot_row, column=1).fill = fill_accent
ws4.cell(row=tot_row, column=1).border = thin_border

for c in range(2, 5):
    cell = ws4.cell(row=tot_row, column=c, value="")
    cell.fill = fill_accent
    cell.border = thin_border

cell_tot = ws4.cell(row=tot_row, column=5, value=f"=SUM(E4:E{tot_row-1})")
cell_tot.font = font_bold
cell_tot.number_format = '$#,##0'
cell_tot.alignment = align_right
cell_tot.fill = fill_accent
cell_tot.border = thin_border

# ---------------------------------------------------------
# Auto-adjust column widths
# ---------------------------------------------------------
for ws in [ws1, ws2, ws3, ws4]:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            # Skip title row length calculation
            if cell.row == 1:
                continue
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

# Save Workbook
output_filename = "AI_x_3D_Oyun_Motoru_Master_Plan.xlsx"
wb.save(output_filename)
print(f"File saved successfully as {output_filename}")