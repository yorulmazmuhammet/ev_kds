import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
from io import BytesIO

# Sayfa ayarları
st.set_page_config(page_title="Karar Destek Sistemi-", layout="wide")
# İngilizce file_uploader metinlerini Türkçeleştiren CSS Bloğu
# İngilizce file_uploader metinlerini Türkçeleştiren Güçlendirilmiş CSS Bloğu
st.markdown("""
    <style>
        /* 1. "Drag and drop file here" metnini gizle ve değiştir */
        div[data-testid="stFileUploadDropzone"] > div > div > span {
            display: none !important;
        }
        div[data-testid="stFileUploadDropzone"] > div > div::before {
            content: "Dosyayı buraya sürükleyip bırakın" !important;
            display: block;
            margin-bottom: 5px;
        }

        /* 2. "Browse files" butonunu şeffaf yap ve üzerine Türkçe metin ekle */
        div[data-testid="stFileUploadDropzone"] button {
            color: transparent !important;
        }
        div[data-testid="stFileUploadDropzone"] button::after {
            content: "Dosyalara Gözat" !important;
            color: #31333F !important; /* Butonun standart koyu gri rengi */
            position: absolute;
            left: 50%;
            top: 50%;
            transform: translate(-50%, -50%);
            visibility: visible;
        }

        /* 3. "Limit 200MB" alt metnini şeffaf yap ve Türkçe limit yaz */
        div[data-testid="stFileUploadDropzone"] small {
            color: transparent !important;
        }
        div[data-testid="stFileUploadDropzone"] small::before {
            content: "Dosya başına boyut sınırı: 200MB" !important;
            color: #888 !important; /* Standart açık gri renk */
            display: block;
            text-align: center;
        }
    </style>
""", unsafe_allow_html=True)


# Literatürde kabul edilen AHP Rassallık İndeksi (Random Index - RI) değerleri (Saaty, 1980)
RI_dict = {1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49, 11: 1.51, 12: 1.53, 13: 1.56, 14: 1.57, 15: 1.59}

st.title("Ortak Konut Şarj İstasyonu Yatırımları Değerlendirilmesine Yönelik Karar Destek Sistemi")
st.markdown("Bu sistem, yüklediğiniz kriterleri otomatik okuyarak dinamik bir şekilde ağırlıkları oluşturur. **Kriterleriniz ve seçenekleriniz dosyadan algılanır.**")

# ==========================================
# TASLAK EXCEL OLUŞTURMA FONKSİYONLARI (DİNAMİK)
# ==========================================
with st.sidebar:
    st.header("📥 Kriter Şablonu İndir")
    st.write("Sistemin kriterlerinizi okuyabilmesi için lütfen kaç kriterle çalışacağınızı seçip şablon dosyasını indirin:")
    
    # Kullanıcıdan kaç kriterlik bir şablon istediğini alıyoruz
    num_template_crit = st.number_input("Kriter Sayısı:", min_value=2, max_value=15, value=7)
    
    # Dinamik kriter isimleri üretme (Örn: "K1: Kriter Açıklaması")
    template_criteria = [f"K{i+1}: Kriter Kısa Açıklaması" for i in range(num_template_crit)]
    template_codes = [f"K{i+1}" for i in range(num_template_crit)]

    def create_ahp_template():
        df_ahp = pd.DataFrame(np.ones((num_template_crit, num_template_crit)), index=template_criteria, columns=template_codes)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_ahp.to_excel(writer, index=True, sheet_name='AHP_Matrisi')
        return output.getvalue()

    def create_site_template():
        df_site = pd.DataFrame(index=["Alternatif A", "Alternatif B", "Alternatif C"], columns=template_codes)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_site.to_excel(writer, index=True, sheet_name='Alternatif_Degerleri')
        return output.getvalue()

    st.download_button(label="Kriter Şablonu İndir", data=create_ahp_template(), file_name="ahp_sablon.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button(label="Alternatifler Şablonu İndir", data=create_site_template(), file_name="alternatif_degerleri_sablon.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# 1. AHP İLE KRİTER AĞIRLIKLARI BÖLÜMÜ
# ==========================================
st.header("1. Kriter Ağırlıklarının Hesaplanması (AHP)")
ahp_file = st.file_uploader("Kriterlerin İkili Karşılaştırma Matrisini İçeren Excel Dosyasını Yükleyin", type=["xlsx"], key="ahp_upload")

# Başlangıçta boş değerler (Dosya yüklenince dolacak)
weights = None
criteria = []
criteria_codes = []
is_benefit = []
n = 0

if ahp_file is not None:
    try:
        # Excel dosyasından AHP matrisini okuma
        df_ahp = pd.read_excel(ahp_file, index_col=0)
        
        # KRİTERLERİ DOSYADAN DİNAMİK OLARAK OKUMA
        criteria = df_ahp.index.astype(str).tolist()
        n = len(criteria)
        # K1, K2 gibi kısa kodları almak için (İki nokta öncesini alır, yoksa tamamını alır)
         
        criteria_codes = [c.split(':')[0].strip() for c in criteria]
        
        
        matrix = df_ahp.values.astype(float)
        st.success(f"İkili karşılaştırma matrisi başarıyla yüklendi! Sistemde **{n} adet kriter** tespit edildi.")
        
        # AHP Hesaplamaları
        col_sums = matrix.sum(axis=0)
        norm_matrix = matrix / col_sums
        weights = norm_matrix.mean(axis=1)

        # Tutarlılık Kontrolü (CR)
        weighted_sum = np.dot(matrix, weights)
        lambda_max = np.mean(weighted_sum / weights)
        CI = (lambda_max - n) / (n - 1)
        RI = RI_dict.get(n, 1.49) # Listede olmayan büyük N değerleri için varsayılan 1.49
        CR = CI / RI if RI != 0 else 0

        col1, col2 = st.columns(2)
        with col1:
            st.write("### Okunan Kriterler ve Ağırlıkları (W)")
            weight_df = pd.DataFrame({"Kriter (Dosyadan Okunan)": criteria, "Ağırlık": weights})
            st.dataframe(weight_df.style.format({"Ağırlık": "{:.2f}"}))

        with col2:
            st.write("### AHP Tutarlılık Kontrolü")
            st.info(f"Hesaplanan Tutarlılık Oranı (CR): **{CR:.4f}**")
            if CR > 0.10:
                st.error("⚠️ DİKKAT: Tutarlılık Oranı (CR) 0.10'dan büyüktür! Lütfen Excel dosyanızdaki ikili karşılaştırma değerlerini gözden geçirin, mantıksal tutarsızlık var.")
            else:
                st.success("✅ BAŞARILI: Tutarlılık Oranı (CR) 0.10'dan küçük veya eşittir. Uzman değerlendirmeleriniz tutarlıdır.")

        # --- DİNAMİK FAYDA / MALİYET SEÇİM EKRANI ---
        st.write("### Kriter Yönlerini Belirleyin")
        st.markdown("MOORA algoritmasının doğru çalışması için dosyadan okunan her bir kriterin **Fayda** (yüksek değer iyidir) veya **Maliyet** (düşük değer iyidir) yönlü olduğunu seçiniz:")
        
        ui_cols = st.columns(3)
        for idx, crit in enumerate(criteria):
            with ui_cols[idx % 3]:
                # Kullanıcıdan yön bilgisini al
                direction = st.selectbox(f"{criteria_codes[idx]} Yönü", options=["Fayda (Maksimizasyon)", "Maliyet (Minimizasyon)"], key=f"dir_{idx}")
                is_benefit.append(True if direction == "Fayda (Maksimizasyon)" else False)

    except Exception as e:
        st.error(f"AHP dosyası okunurken bir hata oluştu. Lütfen formatı kontrol edin. Hata Detayı: {e}")
else:
    st.warning("Kriter ağırlıklarının ve kriterlerin sisteme tanıtılması için lütfen ikili karşılaştırma Excel dosyanızı yükleyin.")

st.divider()

# ------------------------------------------
# 2. MOORA YÖNTEMİ İLE SİTELERİN SIRALANMASI
# ------------------------------------------
st.header("2. Alternatiflerin Sıralanması (MOORA)")
site_file = st.file_uploader("Alternatif İsimleri ve Kriter Değerlerini İçeren Excel Dosyasını Yükleyin", type=["xlsx"], key="site_upload")

if site_file is not None:
    try:
        df_input = pd.read_excel(site_file, index_col=0)
        st.success("Alternatif verileri Excel'den başarıyla yüklendi!")
        st.write("### Yüklenen Alternatif - Kriter Değerleri Tablosu")
        st.dataframe(df_input, use_container_width=True)

        if weights is not None:
            if st.button("MOORA Puanlarını Hesapla ve Sırala", type="primary"):
                # Veride boş hücre varsa uyar
                if df_input.isnull().values.any():
                    st.warning("Excel tablosunda boş hücreler var! Lütfen tüm kriter değerlerini doldurup dosyayı tekrar yükleyin.")
                # Yüklenen site dosyasındaki sütun sayısı ile AHP dosyasındaki kriter sayısı uyuşmuyor ise uyar
                elif len(df_input.columns) != n:
                    st.error(f"Uyuşmazlık Hatası: AHP dosyasında {n} kriter okundu ancak Site dosyasında {len(df_input.columns)} sütun var.")
                else:
                    alt_names_current = df_input.index.tolist()
                    num_alt_current = len(alt_names_current)
                    X = df_input.values.astype(float)
                    
                    # 1. Normalizasyon (MOORA)
                    denom = np.sqrt((X**2).sum(axis=0))
                    denom[denom == 0] = 1.0 
                    X_norm = X / denom
                    
                    # 2. Ağırlıklı Normalize Matris
                    X_weighted = X_norm * weights
                    
                    # 3. Değerlendirme Puanı (yi)
                    y_scores = np.zeros(num_alt_current)
                    for i in range(num_alt_current):
                        sum_benefit = 0.0
                        sum_cost = 0.0
                        for j in range(n):
                            if is_benefit[j]:
                                sum_benefit += X_weighted[i, j]
                            else:
                                sum_cost += X_weighted[i, j]
                        y_scores[i] = sum_benefit - sum_cost
                        
                    # Sonuç DataFrame'i
                    df_result = pd.DataFrame({
                        "Alternatif": alt_names_current,
                        "MOORA Puanı": y_scores
                    })
                    df_result = df_result.sort_values(by="MOORA Puanı", ascending=False).reset_index(drop=True)
                    df_result.index = df_result.index + 1
                    
                    st.write("### Nihai Sıralama Sonuçları")
                    st.dataframe(df_result.style.format({"MOORA Puanı": "{:.2f}"}), use_container_width=True)
                    
                    # Plotly ile Grafik
                    fig = px.bar(
                        df_result, 
                        x="Alternatif", 
                        y="MOORA Puanı", 
                        text="MOORA Puanı", 
                        title="Alternatiflerin MOORA Puanlarına Göre Sıralaması",
                        color="MOORA Puanı",
                        color_continuous_scale="Viridis"
                    )
                    fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.error("MOORA hesaplaması yapılabilmesi için 1. Adımdaki Excel dosyasının yüklenmiş olması gerekmektedir.")
            
    except Exception as e:
        st.error(f"Alternatif dosyası okunurken bir hata oluştu. Lütfen formatı kontrol edin. Hata Detayı: {e}")
else:
    st.warning("Alternatiflerin sıralanabilmesi için lütfen veri dosyasını yükleyin.")