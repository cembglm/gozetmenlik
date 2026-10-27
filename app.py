import streamlit as st
import pandas as pd
import os
import unicodedata

st.set_page_config(page_title="Sınav Görevlendirme Arama", layout="wide")

st.title("📗 Sınav Görevlendirme Arama Sistemi")
st.caption("⚠️ Bu sistem demo amaçlıdır. Lütfen görevlendirme bilgilerinizi resmi listelerden kontrol ediniz.")

name_tab, date_tab = st.tabs(["İsim arama", "Gün / tarih arama"])

with name_tab:
    name = st.text_input("Aranan Bilgi (isim-Soyisim veya ders kodu)", placeholder="Örn: Ahmet Yılmaz veya MAT101", key="name_input")
    search_button = st.button("🔍 Ara", use_container_width=True)

# Dosya yollarını dinamik olarak belirle
base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else '.'
current_file = os.path.join(base_dir, "2026-2027-guz-vize.xlsx")

@st.cache_data
def load_excel(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Dosya bulunamadı: {os.path.basename(path)}")

    try:
        return pd.read_excel(path, sheet_name="OİS Listesi")
    except ValueError:
        # Beklenen sheet adı yoksa ilk sheet'i kullanarak uygulamanın açılmasını garanti et.
        return pd.read_excel(path)

# Excel dosyasını yükle
try:
    df_current = load_excel(current_file)
except Exception as e:
    st.error(f"Veri dosyası yüklenemedi: {e}")
    st.info("Lütfen depoda 2026-2027-guz-vize.xlsx dosyasının bulunduğunu kontrol edin.")
    st.stop()

date_cols = ["Tarih", "Date", "Sınav Tarihi", "Exam Date"]
date_col = next((column for column in df_current.columns if column in date_cols), None)

with date_tab:
    if date_col is None:
        st.warning("Tarihe göre sınavları göstermek için Excel'de bir tarih sütunu bulunamadı.")
    else:
        dates = pd.to_datetime(df_current[date_col], errors="coerce", dayfirst=True).dt.normalize()
        available_dates = sorted(dates.dropna().unique())

        if available_dates:
            day_mapping = {
                "Monday": "Pazartesi",
                "Tuesday": "Salı",
                "Wednesday": "Çarşamba",
                "Thursday": "Perşembe",
                "Friday": "Cuma",
                "Saturday": "Cumartesi",
                "Sunday": "Pazar",
            }
            selected_date = st.selectbox(
                "Sınav tarihini seçin",
                available_dates,
                format_func=lambda date: f"{date.strftime('%d.%m.%Y')} - {day_mapping[date.day_name()]}",
            )
            exams_on_date = df_current.loc[dates == selected_date].copy()
            st.write(f"**{selected_date.strftime('%d.%m.%Y')} tarihindeki sınavlar ({len(exams_on_date)}):**")
            st.dataframe(exams_on_date, use_container_width=True, hide_index=True)
        else:
            st.info("Excel'de görüntülenebilir bir sınav tarihi bulunamadı.")

def normalize_turkish(text):
    """Normalize case and Turkish/Latin accents for name searches."""
    text = unicodedata.normalize("NFKD", str(text).casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return text.replace("ı", "i")

def search(df, name):
    name_normalized = normalize_turkish(name.strip())
    def search_row(row):
        try:
            # NaN değerleri filtreleyip string'e çevir
            row_str = ' '.join([str(x) for x in row if pd.notna(x)])
            return name_normalized in normalize_turkish(row_str)
        except Exception:
            return False
    return df[df.apply(search_row, axis=1)]

# Sadece isim girildiğinde arama yap (Enter veya Buton)
with name_tab:
    if name and name.strip():
        result = search(df_current, name).copy()

        if result.empty:
            st.info("Aranan isimle ilgili güncel dosyada bir sınav görevlendirme kaydı bulunamadı.")
        else:
            # Olası kolon isimleri
            date_cols = ["Tarih", "Date", "Sınav Tarihi", "Exam Date"]
            time_start_cols = ["Başlangıç Saat", "Başlangıç", "Start Time"]
            time_end_cols = ["Bitiş Saat", "Bitiş", "End Time"]

            # Kolonları otomatik bul
            date_col = next((c for c in result.columns if c in date_cols), None)
            time_start_col = next((c for c in result.columns if c in time_start_cols), None)
            time_end_col = next((c for c in result.columns if c in time_end_cols), None)

            if date_col is None:
                st.error("❌ Tarih sütunu bulunamadı. Excel'deki tarih kolon adını söyle, ekleyeyim.")
            else:
                result["__date"] = pd.to_datetime(result[date_col], errors="coerce", dayfirst=True)

                # Zamanı normalize et (Başlangıç-Bitiş formatında)
                if time_start_col and time_end_col:
                    # Başlangıç ve bitiş saatlerini birleştir (örn: "09:00-11:00")
                    result["__time"] = (
                        result[time_start_col].astype(str).str.replace(r'\.\d+', '', regex=True).str[:5] + "-" +
                        result[time_end_col].astype(str).str.replace(r'\.\d+', '', regex=True).str[:5]
                    )
                    # NaN değerleri handle et
                    result["__time"] = result["__time"].replace('nan-nan', 'TBD').replace('NaT-NaT', 'TBD')
                else:
                    result["__time"] = "TBD"

                # Günü normalize et (her zaman tarihten çıkar)
                # Türkçe gün isimleri için manuel çeviri
                day_mapping = {
                    'Monday': 'Pazartesi',
                    'Tuesday': 'Salı',
                    'Wednesday': 'Çarşamba',
                    'Thursday': 'Perşembe',
                    'Friday': 'Cuma',
                    'Saturday': 'Cumartesi',
                    'Sunday': 'Pazar'
                }
                result["__day"] = result["__date"].dt.day_name().map(day_mapping)

                # ------------------------------------------------
                #  📆 TAKVİM GÖRÜNÜMÜ — SAAT x TARİH+GÜN MATRİKSİ
                # ------------------------------------------------
                st.write("## 📆 Takvim Görünümü ")

                # Türkçe ay isimleri
                ay_mapping = {
                    1: 'Ocak', 2: 'Şubat', 3: 'Mart', 4: 'Nisan',
                    5: 'Mayıs', 6: 'Haziran', 7: 'Temmuz', 8: 'Ağustos',
                    9: 'Eylül', 10: 'Ekim', 11: 'Kasım', 12: 'Aralık'
                }

                # Tarih + Gün kombinasyonu oluştur (örn: "26 Ocak (Paz)")
                result["GUN"] = result["__date"].dt.day.astype(str) + " " + result["__date"].dt.month.map(ay_mapping) + " (" + result["__day"].str[:3] + ")"

                # Saat aralığı yoksa tek saat göster
                result["SAAT"] = result["__time"]

                # Derslik sütununu bul
                salon_cols = ["Sınıf", "Salon", "Derslik", "Room", "Classroom"]
                salon_col = next((c for c in result.columns if c in salon_cols), None)

                # Section sütununu düzelt (1.0 -> 1)
                if "Section" in result.columns:
                    result["Section"] = result["Section"].apply(
                        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.','').isdigit() else str(x)
                    )

                # Ders adı, section ve derslik bilgisini birleştir
                if salon_col and "Section" in result.columns:
                    result["TAKVIM_BILGI"] = (
                        result["Ders Kodu"].astype(str) +
                        " (Sec:" + result["Section"].astype(str) + ") " +
                        result[salon_col].astype(str)
                    )
                elif salon_col:
                    result["TAKVIM_BILGI"] = result["Ders Kodu"].astype(str) + " (" + result[salon_col].astype(str) + ")"
                else:
                    result["TAKVIM_BILGI"] = result["Ders Kodu"].astype(str)

                # Takvim tablosu oluştur
                takvim = result.pivot_table(
                    index="SAAT",
                    columns="GUN",
                    values="TAKVIM_BILGI",
                    aggfunc=lambda x: " | ".join(x.astype(str))
                )

                # Sütunları tarihe göre sırala
                sorted_columns = sorted(takvim.columns, key=lambda x: result[result["GUN"] == x]["__date"].iloc[0])
                takvim = takvim.reindex(columns=sorted_columns, fill_value="")

                # NaN değerleri boş string'e çevir
                takvim = takvim.fillna("")

                # Tamamen boş satırları kaldır (tüm sütunlar boş olan satırlar)
                takvim = takvim[(takvim != "").any(axis=1)]

                # Toplam sınav sayısını hesapla
                total_exams = (takvim != "").sum().sum()
                st.write(f"**Toplam {total_exams} adet sınav görevlendirmesi bulundu.**")

                # Takvim tablosunu renklendir ve genişlik ayarları yap
                styled_takvim = takvim.style\
                    .set_properties(**{
                        'white-space': 'pre-wrap',
                        'text-align': 'left',
                        'font-size': '11px',
                        'padding': '8px'
                    })

                # Dinamik yükseklik hesapla (satır sayısı * 40px + başlık için 40px)
                dynamic_height = len(takvim) * 40 + 40

                st.dataframe(
                    styled_takvim,
                    use_container_width=True,
                    height=dynamic_height  # Dinamik takvim yüksekliği
                )

                # ------------------------------------------------
                #  📋 DETAYLI LİSTE GÖRÜNÜMÜ
                # ------------------------------------------------
                st.write("## 📋 Detaylı Liste Görünümü")

                # Tarihe ve saate göre sırala
                result_sorted = result.sort_values(by=["__date", "__time"])

                # Gereksiz kolonları gizle
                cols_to_hide = ["__date", "__time", "__day", "GUN", "SAAT", "TAKVIM_BILGI"]
                display_df = result_sorted.drop(columns=[c for c in cols_to_hide if c in result_sorted.columns])

                # Gözetmen sütunlarını vurgula
                def highlight_searched_name(row):
                    name_norm = normalize_turkish(name)
                    styles = [''] * len(row)

                    # Gözetmen sütunlarını kontrol et
                    gozetmen_cols = [c for c in row.index if 'Gözetmen' in str(c)]
                    for idx, col in enumerate(row.index):
                        if col in gozetmen_cols:
                            cell_value = str(row[col])
                            if name_norm in normalize_turkish(cell_value):
                                styles[idx] = 'background-color: #ffffcc; font-weight: bold'

                    return styles

                styled_df = display_df.style.apply(highlight_searched_name, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                # ------------------------------------------------
                #  📊 İSTATİSTİKLER
                # ------------------------------------------------
                st.write("## 📊 İstatistikler")

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric("Toplam Görevlendirme", len(result_sorted))

                with col2:
                    # Benzersiz dersleri say
                    unique_courses = result_sorted['Ders Kodu'].nunique()
                    st.metric("Farklı Ders Sayısı", unique_courses)

                with col3:
                    # Benzersiz günleri say
                    unique_days = result_sorted['__date'].nunique()
                    st.metric("Farklı Gün Sayısı", unique_days)

    else:
        st.info("👆 Lütfen aramak istediğiniz bilgiyi girin.")
