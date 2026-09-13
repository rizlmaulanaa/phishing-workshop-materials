#!/usr/bin/env python3
"""
Generate PowerPoint slides for Phishing Simulation Workshop
Requires: pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def add_title_slide(prs, title, subtitle):
    """Add a title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(15, 15, 15)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 107, 107)
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.7), Inches(9), Inches(1.2))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(32)
    p.font.color.rgb = RGBColor(78, 205, 196)
    
    return slide

def add_content_slide(prs, title, content_items, is_bullet=True):
    """Add a content slide with bullet points or text"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(15, 15, 15)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 107, 107)
    
    # Content
    content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(8.6), Inches(5.5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    for idx, item in enumerate(content_items):
        if idx > 0:
            text_frame.add_paragraph()
        p = text_frame.paragraphs[idx]
        p.text = item
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(200, 200, 200)
        p.level = 0
        p.space_before = Pt(8)
        p.space_after = Pt(8)
        
        if is_bullet:
            p.level = 0
    
    return slide

def create_workshop_slides():
    """Create the complete workshop presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # SLIDE 1: TITLE
    add_title_slide(prs, "🎣 Hands-On Phishing Simulation", 
                    "Workshop 3 Jam - Pemahaman Mendalam tentang Serangan Phishing\n\n"
                    "Dari konsep dasar hingga praktik langsung dengan GoPhish")
    
    # SLIDE 2: AGENDA
    add_content_slide(prs, "📋 Agenda Workshop (180 Menit)", [
        "00:00–00:10  |  Pembukaan & Ice Breaker (10 min)",
        "00:10–00:50  |  Sesi 1: Definisi, Tipe & Dampak Phishing (40 min)",
        "00:50–01:35  |  Sesi 2: Anatomi Serangan Phishing (45 min)",
        "01:35–01:45  |  ☕ Break (10 min)",
        "01:45–03:00  |  Sesi 3: Hands-On GoPhish Simulation (75 min)"
    ])
    
    # SLIDE 3: TUJUAN
    add_content_slide(prs, "🎯 Tujuan Workshop", [
        "✓ Memahami anatomi serangan phishing secara menyeluruh",
        "✓ Mengenali teknik social engineering yang digunakan penyerang",
        "✓ Praktik langsung menjalankan simulasi phishing dengan GoPhish",
        "✓ Merancang kampanye simulasi yang etis dan efektif",
        "✓ Mengubah data simulasi menjadi program pelatihan yang berdampak"
    ])
    
    # SESI 1 INTRO
    add_title_slide(prs, "📚 SESI 1", "Definisi, Tipe & Dampak Phishing\n⏱️ 40 Menit")
    
    # SLIDE 5: APA ITU PHISHING
    add_content_slide(prs, "🎣 Apa Itu Phishing?", [
        "Phishing adalah kelas serangan SOCIAL ENGINEERING di mana penyerang menyamar sebagai entitas terpercaya untuk menipu target agar mengungkapkan kredensial, informasi keuangan, atau data sensitif.",
        "",
        "📌 Istilah 'Phishing' berasal dari 'Fishing'",
        "→ Mekanisme umpan-dan-kail pada PSIKOLOGI MANUSIA, bukan kerentanan teknis",
        "",
        "🔗 Persinggungan Tiga Disiplin:",
        "→ Keamanan Siber + Psikologi Perilaku + Komunikasi Jaringan"
    ])
    
    # SLIDE 6: FILOSOFI
    add_content_slide(prs, "💭 Hack Manusia, Bukan Komputer", [
        "⚙️ HACK TRADISIONAL",
        "   • Target: Kelemahan sistem",
        "   • Metode: Eksploitasi teknis",
        "   • Hasil: Bypass pertahanan",
        "",
        "🧠 PHISHING",
        "   • Target: Kelemahan keputusan",
        "   • Metode: Manipulasi psikologis",
        "   • Hasil: Membujuk pemilik kunci"
    ])
    
    # SLIDE 7: TIPE PHISHING
    add_content_slide(prs, "🎯 Tipe-Tipe Phishing", [
        "Bulk Phishing → Email massal generik ('Akun Anda akan ditutup dalam 24 jam')",
        "",
        "Spear Phishing → Target individu spesifik dengan detail personal",
        "",
        "Whaling → Target eksekutif senior ('CEO meminta transfer')",
        "",
        "Smishing → Phishing via SMS ('Paket tertahan, klik untuk jadwal')",
        "",
        "Vishing → Phishing via telepon suara (penelepon minta OTP)"
    ])
    
    # SLIDE 8: STATISTIK
    add_content_slide(prs, "📊 Dampak Phishing: Data & Statistik", [
        "📈 PREVALENSI",
        "   • 80%+ pelanggaran social engineering melibatkan phishing",
        "   • 70% di antaranya sangat dipersonalisasi",
        "   • 38% bisnis di Inggris alami serangan phishing/tahun",
        "",
        "💰 KERUGIAN FINANSIAL",
        "   • Phishing/Spoofing: $215.8 Juta (2025)",
        "   • Business Email Compromise: >$3 Miliar",
        "",
        "⚡ KECEPATAN RESPONS: ~21 detik untuk klik"
    ])
    
    # SLIDE 9: MENGAPA SIMULASI
    add_content_slide(prs, "🛡️ Mengapa Simulasi Phishing Penting?", [
        "📌 Phishing Simulation: Email phishing palsu untuk mengukur dan melatih kesadaran karyawan.",
        "",
        "✓ TUJUAN UTAMA: Menciptakan momen pembelajaran, BUKAN menangkap orang",
        "",
        "💡 Penelitian menunjukkan:",
        "Ketika simulasi dirancang untuk menangkap orang, itu menjadi latihan pengawasan. Staf merasa malu dan tidak mendapat kesadaran keamanan yang lebih baik.",
        "",
        "🎓 Kunci Efektivitas: Halaman edukasi yang muncul SEGERA SETELAH KLIK adalah pelatihan yang sebenarnya."
    ])
    
    # SESI 2 INTRO
    add_title_slide(prs, "🔍 SESI 2", "Anatomi Serangan Phishing\n⏱️ 45 Menit")
    
    # SLIDE 11: ENAM TAHAP
    add_content_slide(prs, "⚙️ Enam Tahap Anatomi Serangan", [
        "1️⃣ TARGET → Siapa yang dipilih? (massal atau spesifik)",
        "",
        "2️⃣ LURE (Umpan) → Alasan bertindak? (faktur, pengiriman, password)",
        "",
        "3️⃣ HOOK (Kail) → Pemicu emosional? (urgensi, ketakutan, otoritas)",
        "",
        "4️⃣ LANDING → Ke mana klik diarahkan? (halaman palsu, proxy, lampiran)",
        "",
        "5️⃣ CAPTURE → Apa yang didapat? (kredensial, malware, pembayaran)",
        "",
        "6️⃣ CASH-OUT → Hasil akhir? (akses, ransomware, transfer uang)"
    ])
    
    # SLIDE 12: TEKNIK SOCIAL ENGINEERING
    add_content_slide(prs, "🧩 Teknik Social Engineering: Elemen Kunci", [
        "🔔 URGENCY (Urgensi) → Tekanan waktu agar korban tidak berpikir kritis",
        "",
        "👮 AUTHORITY (Otoritas) → Meniru figur berwenang yang cenderung dituruti",
        "",
        "😨 FEAR (Ketakutan) → Memicu kecemasan akan konsekuensi negatif",
        "",
        "👥 SOCIAL PROOF → Menyarankan orang lain sudah bertindak",
        "",
        "👋 FAMILIARITY (Familiaritas) → Menggunakan merek/kolega terpercaya",
        "",
        "🎁 REWARD/CURIOSITY → Menawarkan hadiah atau memicu penasaran"
    ])
    
    # SLIDE 13: LATIHAN KELOMPOK
    add_content_slide(prs, "👥 Latihan Kelompok: Analisis Email Phishing", [
        "⏱️ Durasi: 15 menit (3-4 orang per kelompok)",
        "",
        "📋 TUGAS: Diberikan satu contoh email phishing",
        "",
        "❓ JAWAB:",
        "1. Siapa target yang dipilih? (Tahap 1)",
        "2. Apa lure/umpan yang digunakan? (Tahap 2)",
        "3. Pemicu emosional apa yang digunakan? (Tahap 3)",
        "4. Ke mana klik akan diarahkan? (Tahap 4)",
        "5. Apa yang penyerang ingin dapatkan? (Tahap 5-6)",
        "",
        "🎯 Diskusikan hasil bersama di kelas"
    ])
    
    # SESI 3 INTRO
    add_title_slide(prs, "💻 SESI 3", "Hands-On GoPhish Simulation\n⏱️ 75 Menit")
    
    # SLIDE 15: GOPHISH INTRO
    add_content_slide(prs, "🎣 Apa Itu GoPhish?", [
        "Framework simulasi phishing OPEN-SOURCE untuk menjalankan kampanye kesadaran phishing yang terotorisasi.",
        "",
        "📧 SENDING PROFILE → Konfigurasi SMTP untuk mengirim email",
        "",
        "✏️ EMAIL TEMPLATE → HTML email yang akan dikirim ke target",
        "",
        "🌐 LANDING PAGE → Halaman yang ditampilkan saat target mengklik tautan",
        "",
        "👥 USERS & GROUPS → Daftar target yang menerima email"
    ])
    
    # SLIDE 16: SETUP ENVIRONMENT
    add_content_slide(prs, "⚙️ Persiapan Lingkungan Lab", [
        "⚠️ PERINGATAN ETIKA:",
        "Simulasi phishing HANYA boleh dijalankan terhadap akun yang Anda kontrol sendiri, atau dengan izin tertulis eksplisit. Gunakan lingkungan lab TERISOLASI.",
        "",
        "OPSI A: Instalasi Langsung (Linux/Kali)",
        "→ cd /opt && wget gophish-v0.12.1-linux-64bit.zip",
        "→ unzip && cd gophish && chmod +x gophish && ./gophish",
        "",
        "Admin panel: https://0.0.0.0:3333 (password di terminal)"
    ])
    
    # SLIDE 17: DOCKER
    add_content_slide(prs, "🐳 Setup Rekomendasi: Docker", [
        "Lebih mudah dan terisolasi. Gunakan MailHog sebagai SMTP sink agar email AMAN (tidak benar-benar terkirim).",
        "",
        "# Jalankan GoPhish",
        "docker run -d --name gophish -p 3333:3333 -p 80:80 \\",
        "  -v ~/gophish-data:/data gophish/gophish:latest",
        "",
        "# Jalankan MailHog (SMTP Sink)",
        "docker run -d --name mailhog -p 1025:1025 -p 8025:8025 mailhog/mailhog",
        "",
        "📧 Email tertangkap di: http://<host>:8025"
    ])
    
    # SLIDE 18: STEP 1
    add_content_slide(prs, "📧 Langkah 1: Konfigurasi Sending Profile (SMTP)", [
        "Navigasi: GoPhish Admin → Sending Profiles → New Profile",
        "",
        "FIELD KONFIGURASI:",
        "   Name: MailHog Local (atau nama yang Anda pilih)",
        "   From: it-support@example.com",
        "   Host: 127.0.0.1:1025 (MailHog) atau smtp-relay.brevo.com:587 (Produksi)",
        "   Username: (kosong untuk MailHog) atau email Anda",
        "   Password: (kosong untuk MailHog) atau API Key",
        "",
        "✓ Klik: 'Send Test Email' untuk verifikasi konfigurasi"
    ])
    
    # SLIDE 19: STEP 2
    add_content_slide(prs, "✏️ Langkah 2: Membuat Email Template", [
        "Navigasi: Email Templates → New Template",
        "",
        "KONFIGURASI:",
        "   Name: Berikan nama template (misal 'IT Security Alert')",
        "   Subject: 'Peringatan Keamanan: Verifikasi Akun Diperlukan'",
        "",
        "KONTEN HTML:",
        "   Gunakan {{.FirstName}} untuk nama korban (variable GoPhish)",
        "   Sertakan {{.URL}} pada tautan - ini akan di-replace dengan landing page URL",
        "",
        "💡 PENTING: {{.URL}} adalah variabel GoPhish yang otomatis mengarahkan ke landing page Anda"
    ])
    
    # SLIDE 20: STEP 3
    add_content_slide(prs, "🌐 Langkah 3: Membuat Landing Page", [
        "Navigasi: Landing Pages → New Page",
        "",
        "DUA OPSI:",
        "   1. MANUAL: Tulis HTML login palsu sendiri",
        "   2. IMPORT: Masukkan URL halaman login asli untuk ditiru otomatis",
        "",
        "SETTING PENTING:",
        "   ✓ Enable 'Capture Submitted Data'",
        "   ✓ Enable 'Capture Passwords'",
        "   ✓ Set 'Redirect to' → URL halaman edukasi",
        "",
        "🎯 KUNCI: Halaman edukasi yang muncul setelah klik adalah PELATIHAN yang sebenarnya"
    ])
    
    # SLIDE 21: STEP 4
    add_content_slide(prs, "👥 Langkah 4: Membuat Target Group", [
        "Navigasi: Users & Groups → New Group",
        "",
        "OPSI 1: Manual Entry",
        "   Budi Santoso | budi@example.com | Staff",
        "   Siti Rahayu | siti@example.com | Manager",
        "",
        "OPSI 2: Import dari CSV",
        "   First Name,Last Name,Email,Position",
        "   Budi,Santoso,budi@example.com,Staff",
        "",
        "⚠️ PENTING: Hanya gunakan email yang ANDA KONTROL SENDIRI atau dengan izin TERTULIS EKSPLISIT"
    ])
    
    # SLIDE 22: STEP 5
    add_content_slide(prs, "🚀 Langkah 5: Meluncurkan Kampanye", [
        "Navigasi: Campaigns → New Campaign",
        "",
        "KONFIGURASI KAMPANYE:",
        "   Name: Workshop Phishing Simulation #1",
        "   Email Template: IT Security Alert",
        "   Landing Page: Fake Login Page",
        "   URL: (Biarkan kosong - menggunakan default)",
        "   Sending Profile: MailHog Local",
        "   Groups: Workshop Test Users",
        "   Launch Date: Now",
        "",
        "✓ Klik: 'Launch Campaign'"
    ])
    
    # SLIDE 23: STEP 6
    add_content_slide(prs, "📊 Langkah 6: Memantau & Menganalisis Hasil", [
        "Pantau dashboard GoPhish secara REAL-TIME",
        "",
        "METRIK YANG TERSEDIA:",
        "   📬 Email Sent → Berapa email berhasil terkirim?",
        "   👁️ Email Opened → Berapa yang dibuka? (via tracking pixel)",
        "   🔗 Link Clicked → Berapa yang mengklik tautan?",
        "   ✍️ Data Submitted → Berapa yang mengisi formulir?",
        "",
        "💡 KPI KESADARAN: Semakin tinggi click rate, semakin besar kebutuhan pelatihan tambahan"
    ])
    
    # SLIDE 24: METRICS TO TRAINING
    add_content_slide(prs, "📈 Dari Metrik ke Program Pelatihan", [
        "Gunakan data simulasi untuk merancang program pelatihan yang TEPAT SASARAN",
        "",
        "Click Rate < 5% → Kesadaran tinggi → Kursus refresh kuartalan",
        "",
        "Click Rate 5-15% → Kesadaran moderat → Micro-learning yang ditargetkan",
        "",
        "Click Rate > 15% → Kesadaran rendah → Pelatihan wajib + komunikasi manajer",
        "",
        "🎓 PRINSIP UTAMA: Simulasi phishing adalah tentang PENDIDIKAN, BUKAN HUKUMAN"
    ])
    
    # BEST PRACTICES
    add_content_slide(prs, "✅ Best Practices: Simulasi yang Etis & Efektif", [
        "1️⃣ RANCANG SKENARIO DULU",
        "   Tanyakan: 'Keterampilan apa yang ingin dikembangkan?' bukan 'Siapa yang tangkap?'",
        "",
        "2️⃣ KALIBRASI KESULITAN",
        "   Mulai dengan skenario yang mudah, tingkatkan kompleksitas seiring kemampuan",
        "",
        "3️⃣ HALAMAN EDUKASI ADALAH PELATIHAN",
        "   Momen setelah klik adalah pembelajaran sebenarnya, bukan laporan 3 hari kemudian",
        "",
        "4️⃣ SESUAIKAN DENGAN PERAN",
        "   Staf yang terima email eksternal ≠ Eksekutif yang ditarget whaling"
    ])
    
    # ETHICAL WARNINGS
    add_content_slide(prs, "⚠️ Peringatan Etika", [
        "❌ JANGAN:",
        "   • Jalankan simulasi tanpa izin TERTULIS dari manajemen & peserta",
        "   • Jalankan di lingkungan PRODUKSI - gunakan VM/lab terisolasi",
        "   • Kumpulkan KREDENSIAL ASLI - hanya catat apakah form dikirim",
        "   • Gunakan sebagai alat HUKUMAN - itu akan merusak kepercayaan",
        "   • Sebarkan hasil ke publik tanpa ANONYMISASI",
        "",
        "✓ LAKUKAN:",
        "   Konsultasi dengan HR & Legal, jelas tentang tujuan pembelajaran, fokus pada pengembangan, hormati privasi peserta"
    ])
    
    # SUMMARY
    add_content_slide(prs, "📚 Ringkasan & Sumber Daya", [
        "🎯 POIN KUNCI:",
        "   ✓ Phishing menargetkan MANUSIA, bukan teknologi",
        "   ✓ Pola: Target → Lure → Hook → Landing → Capture → Cash-out",
        "   ✓ Teknik: Urgensi, Otoritas, Ketakutan, Social Proof, Familiaritas",
        "   ✓ GoPhish memungkinkan simulasi TERKONTROL & TERUKUR",
        "   ✓ Efektivitas: PEMBELAJARAN > HUKUMAN",
        "",
        "🔗 REFERENSI:",
        "   • GoPhish: https://getgophish.com",
        "   • Template: https://github.com/HailBytes/gophish-training-templates",
        "   • Verizon DBIR 2025 • Proofpoint Threat Report"
    ])
    
    # FINAL SLIDE
    add_title_slide(prs, "❓ Pertanyaan & Diskusi", 
                    "Mari kita terapkan pengetahuan ini bersama!\n\n"
                    "📧 Hubungi untuk resources tambahan\n"
                    "🔗 GoPhish Docs: https://docs.getgophish.com")
    
    return prs

if __name__ == "__main__":
    prs = create_workshop_slides()
    prs.save("Phishing_Simulation_Workshop.pptx")
    print("✓ PowerPoint presentation created: Phishing_Simulation_Workshop.pptx")
