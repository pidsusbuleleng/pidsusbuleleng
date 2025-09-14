const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
// Fungsi untuk mengubah format tanggal ke format Indonesia
function formatDateID(dateString) {
    if (!dateString) return '';
    const [year, month, day] = dateString.split('-');
    const monthNames = [
        'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
        'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    const monthName = monthNames[parseInt(month, 10) - 1];
    return `${parseInt(day, 10)} ${monthName} ${year}`;
}

const app = express();
app.use(express.static('Asset'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());


// Konfigurasi utama yang mencakup tahapan dan template di dalamnya
const configs = {
    'penyelidikan': {
        label: 'Penyelidikan',
        templates: {
            'surat1_lidik': {
                file: 'Pidsus-5A.docx',
                label: 'Pidsus-5A',
                fields: [
                    { name: 'nama', label: 'Nama', type: 'text' },
                    { name: 'jabatan', label: 'Jabatan', type: 'text' }
                ]
            },
            'surat2_lidik': {
                file: 'Pidsus-5B.docx',
                label: 'Pidsus-5B',
                fields: [
                    { name: 'nama', label: 'Nama', type: 'text' },
                    { name: 'jabatan', label: 'Jabatan', type: 'text' },
                    { name: 'jabatan', label: 'Jabatan', type: 'text' },
                    { name: 'jabatan', label: 'Jabatan', type: 'text' }
                ]
            }
        }
    },
    'penyidikan_umum': {
        label: 'Penyidikan Umum',
        templates: {
            'surat2_dikum': {
                file: 'surat2.docx',
                label: 'Surat 2 (Penyidikan Umum)',
                fields: [
                    { name: 'nama', label: 'Nama', type: 'text' },
                    { name: 'jabatan', label: 'Jabatan', type: 'text' }
                ]
            }
        }
    },
    'penyidikan_khusus': {
        label: 'Penyidikan Khusus',
        templates: {
            'surat3_diksus': {
                file: 'surat3.docx',
                label: 'Surat 3 (Penyidikan Khusus)',
                fields: [
                    { name: 'nama', label: 'Nama', type: 'text' },
                    { name: 'nomor_kasus', label: 'Nomor Kasus', type: 'text' }
                ]
            }
        }
    },
    'penuntutan': {
        label: 'Penuntutan',
        templates: {
            'P-36 BRI': {
                file: 'P-36 BRI.docx',
                label: 'P-36 BRI (Semua Terdakwa)',
                fields: [
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' },
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' }
                ]
            },
            'P-37 GEDE GAWATRA': {
                file: 'P-37 GEDE GAWATRA.docx',
                label: 'P-37 GEDE GAWATRA',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' }
                ]
            },
            'P-37 I MADE DWI MEI ANGGARA': {
                file: 'P-37 I MADE DWI MEI ANGGARA.docx',
                label: 'P-37 I MADE DWI MEI ANGGARA',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' }
                ]
            },
            'P-37 WAYAN EDI SUPARMAN': {
                file: 'P-37 WAYAN EDI SUPARMAN.docx',
                label: 'P-37 WAYAN EDI SUPARMAN',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' }
                ]
            },
            'P-37 SAKSI BRI': {
                file: 'P-37 SAKSI BRI.docx',
                label: 'P-37 SAKSI BRI',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'namasaksi', label: 'Nama Saksi', type: 'text' },
                    { name: 'tempatlahir', label: 'Tempat Lahir', type: 'text' },
                    { name: 'umur', label: 'Umur', type: 'number' },
                    { name: 'tgllahir', label: 'Tanggal Lahir', type: 'date' },
                    { name: 'jeniskelamin', label: 'Jenis Kelamin', type: 'text' },
                    { name: 'tempattinggal', label: 'Alamat', type: 'text' },
                    { name: 'agama', label: 'Agama', type: 'text' },
                    { name: 'pekerjaan', label: 'Pekerjaan', type: 'text' },
                    { name: 'pendidikan', label: 'Pendidikan', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' }
                ]
            },
            'P-38 GEDE GAWATRA': {
                file: 'P-38 GEDE GAWATRA.docx',
                label: 'P-38 GEDE GAWATRA',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' }
                ]
            },
            'P-38 I MADE DWI MEI ANGGARA': {
                file: 'P-38 I MADE DWI MEI ANGGARA.docx',
                label: 'P-38 I MADE DWI MEI ANGGARA',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' }
                ]
            },
             'P-38 WAYAN EDI SUPARMAN': {
                file: 'P-38 WAYAN EDI SUPARMAN.docx',
                label: 'P-38 WAYAN EDI SUPARMAN',
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' }
                ]
            },
            'P-38 SAKSI BRI': {
                file: 'P-38 SAKSI BRI.docx',
                label: 'P-38 SAKSI BRI',
                hasDynamicTable: true,
                fields: [
                    { name: 'nomorsurat', label: 'Nomor Surat', type: 'text' },
                    { name: 'tglsurat', label: 'Tanggal Surat', type: 'date' },
                    { name: 'kepada', label: 'Kepada', type: 'text' },
                    { name: 'di', label: 'Di', type: 'text' },
                    { name: 'tglsidang', label: 'Tanggal Persidangan', type: 'date' },
                    { name: 'harisidang', label: 'Hari Persidangan', type: 'text' },
                    { name: 'jamsidang', label: 'Jam Persidangan', type: 'text' }
                ]
            }
        }
    }
};

// Halaman utama: Pilih Tahapan Perkara
app.get('/', (req, res) => {
    const stageKeys = Object.keys(configs);
    res.send(`
        <html>
        <head>
            <title>Shortcut Persuratan Pidsus Buleleng</title>
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <style>
                html, body {
                    overflow: hidden; /* Mencegah scroll page secara default */
                    height: 100%;
                }
                body {
                    background-image: url('https://raw.githubusercontent.com/pidsusbuleleng/pidsusbuleleng/refs/heads/main/Asset/back.jpg');
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                    background-color: #2c2c2c; /* Fallback color */
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    display: flex;
                    flex-direction: column;
                    min-height: 100vh;
                    color: #fff;
                }
                .header {
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    padding: 20px;
                    background: rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(15px);
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    z-index: 1000;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.2);
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    height: 64px; /* Menetapkan tinggi eksplisit */
                    box-sizing: border-box; /* Memastikan padding termasuk dalam tinggi */
                }
                .header-title {
                    font-size: 24px;
                    font-weight: 700;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                    margin: 0;
                    line-height: 1; /* Menyamakan line-height */
                }
                .nav-buttons {
                    display: flex;
                    gap: 10px;
                }
                .button-link {
                    padding: 12px 20px;
                    border-radius: 12px;
                    border: none;
                    background: rgba(255, 255, 255, 0.15);
                    color: #fff;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: background 0.3s, transform 0.3s;
                    backdrop-filter: blur(10px);
                    border: 1px solid rgba(255,255,255,0.3);
                    text-decoration: none;
                    display: inline-block;
                    text-align: center;
                }
                .button-link:hover {
                    background: rgba(255, 255, 255, 0.25);
                    transform: translateY(-2px);
                }
                main {
                    flex: 1;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    padding: 120px;
                    box-sizing: border-box;
                    position: relative;
                    z-index: 1;
                    overflow: hidden;
                }
                main::before {
                    content: '';
                    position: absolute;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: linear-gradient(135deg, rgba(0, 0, 0, 0.3) 0%, rgba(0, 0, 0, 0.2) 100%);
                    z-index: -1;
                }
                .container {
                    max-width: 400px;
                    width: 100%;
                    background: rgba(255, 255, 255, 0.1);
                    border-radius: 25px;
                    box-shadow: 0 10px 40px rgba(0,0,0,0.1), 0 0 0 1px rgba(255,255,255,0.2);
                    padding: 35px 30px;
                    box-sizing: border-box;
                    backdrop-filter: blur(35px);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                    margin-top: 0; /* Tambahkan margin agar tidak tertutup header */
                }
                h2 {
                    text-align: center;
                    color: #fff;
                    font-weight: 700;
                    margin-bottom: 30px;
                    font-size: 32px;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                }
                ul { list-style: none; padding: 0; margin: 0; }
                li {
                    padding: 16px 20px;
                    margin-bottom: 15px;
                    background: rgba(255, 255, 255, 0.15);
                    border-radius: 18px;
                    transition: background 0.3s, transform 0.3s, box-shadow 0.3s;
                    text-align: center;
                    border: 1px solid rgba(255,255,255,0.3);
                    backdrop-filter: blur(10px);
                }
                li:last-child { margin-bottom: 0; }
                li:hover {
                    background: rgba(255, 255, 255, 0.25);
                    transform: translateY(-3px);
                    box-shadow: 0 8px 20px rgba(0,0,0,0.12);
                    cursor: pointer;
                }
                li a {
                    text-decoration: none;
                    color: #fff;
                    font-weight: 500;
                    display: block;
                    font-size: 17px;
                }
                .footer {
                    flex-shrink: 0; /* Mencegah footer menyusut */
                    padding: 20px 0;
                    text-align: center;
                    color: rgba(255, 255, 255, 0.6);
                    font-size: 14px;
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(5px);
                    border-top: 1px solid rgba(255, 255, 255, 0.1);
                    position: relative; /* Ubah dari fixed ke relative */
                }
                .form-container {
                    display: flex;
                    flex-direction: column;
                    gap: 15px; /* Memberi jarak antar elemen form */
                }

                @media (max-width: 600px) {
                    .form-container {
                        width: 90%; /* Buat lebar form 90% dari layar */
                        padding: 10px;
                    }
                }
            </style>
        </head>
        <body>
            <header class="header">
                <h1 class="header-title">Shortcut Persuratan Pidsus Buleleng</h1>
            </header>
            <main>
                <div class="container">
                    <h2>Pilih Tahapan Perkara</h2>
                    <ul>
                        ${stageKeys.map(key => `
                            <li><a href="/templates?stage=${encodeURIComponent(key)}">${configs[key].label}</a></li>
                        `).join('')}
                    </ul>
                </div>
            </main>
            <div class="footer">
                <p>Hak Cipta © parwatabayu. Pidsus 2025.</p>
            </div>
        </body>
        </html>
    `);
});

// Halaman: Pilih Template Surat (berdasarkan tahapan yang dipilih)
app.get('/templates', (req, res) => {
    const stage = req.query.stage;
    const config = configs[stage];

    if (!config) {
        return res.redirect('/');
    }

    const templates = config.templates;
    res.send(`
        <html>
        <head>
            <title>Pilih Template Surat</title>
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
            <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
            <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
            <style>
                html, body {
                    overflow: hidden;
                    height: 100%;
                }
                body {
                    background-image: url('https://raw.githubusercontent.com/pidsusbuleleng/pidsusbuleleng/refs/heads/main/Asset/back.jpg');
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                    background-color: #2c2c2c;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    display: flex;
                    flex-direction: column;
                    min-height: 100vh;
                    color: #fff;
                }
                .header {
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    padding: 20px;
                    background: rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(15px);
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    z-index: 1000;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.2);
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    height: 64px; /* Menetapkan tinggi eksplisit */
                    box-sizing: border-box; /* Memastikan padding termasuk dalam tinggi */
                }
                .header-title {
                    font-size: 24px;
                    font-weight: 700;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                    margin: 0;
                    line-height: 1; /* Menyamakan line-height */
                }
                .nav-buttons {
                    display: flex;
                    gap: 10px;
                }
                .button-link {
                    padding: 12px 20px;
                    border-radius: 12px;
                    border: none;
                    background: rgba(255, 255, 255, 0.15);
                    color: #fff;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: background 0.3s, transform 0.3s;
                    backdrop-filter: blur(10px);
                    border: 1px solid rgba(255,255,255,0.3);
                    text-decoration: none;
                    display: inline-block;
                    text-align: center;
                }
                .button-link:hover {
                    background: rgba(255, 255, 255, 0.25);
                    transform: translateY(-2px);
                }
                main {
                    flex: 1;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    padding: 120px;
                    box-sizing: border-box;
                    position: relative;
                    z-index: 1;
                    overflow-y: auto; /* Tambahkan ini agar bisa di-scroll jika kontennya panjang */
                }
                main::before {
                    content: '';
                    position: absolute;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: linear-gradient(135deg, rgba(0, 0, 0, 0.3) 0%, rgba(0, 0, 0, 0.2) 100%);
                    z-index: -1;
                }
                .container {
                    max-width: 450px;
                    width: 100%;
                    background: rgba(255, 255, 255, 0.1);
                    border-radius: 25px;
                    box-shadow: 0 10px 40px rgba(0,0,0,0.1), 0 0 0 1px rgba(255,255,255,0.2);
                    padding: 35px 30px;
                    box-sizing: border-box;
                    backdrop-filter: blur(35px);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                    margin-top: 0;
                }
                h2 {
                    text-align: center;
                    color: #fff;
                    font-weight: 700;
                    margin-bottom: 30px;
                    font-size: 32px;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                }
                .form-group {
                    display: flex;
                    flex-direction: column;
                    gap: 15px;
                    margin-bottom: 25px;
                }
                .select2-container { flex-grow: 1; width: 100% !important; }
                .select2-container .select2-selection--single {
                    height: 50px;
                    border: 1px solid rgba(255, 255, 255, 0.4);
                    border-radius: 18px;
                    background: rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(10px);
                    transition: all 0.2s;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                }
                .select2-container--default .select2-selection--single .select2-selection__rendered {
                    line-height: 48px;
                    color: #fff;
                    padding-left: 20px;
                    font-size: 16px;
                }
                .select2-container--default .select2-selection--single .select2-selection__arrow {
                    height: 48px;
                    right: 12px;
                }
                .select2-container--default.select2-container--focus .select2-selection--single {
                    border: 1.5px solid #8EC5FC;
                    box-shadow: 0 0 0 4px rgba(142, 197, 252, 0.2);
                }
                .select2-dropdown {
                    border-radius: 18px;
                    border: 1px solid rgba(255, 255, 255, 0.4);
                    box-shadow: 0 8px 25px rgba(0,0,0,0.15);
                    background: rgba(255, 255, 255, 0.15);
                    backdrop-filter: blur(25px);
                    overflow: hidden;
                    position: absolute;
                    z-index: 1051;
                }
                .select2-search--dropdown .select2-search__field {
                    border-radius: 15px;
                    border: 1px solid rgba(255,255,255,0.4);
                    background: rgba(255, 255, 255, 0.1);
                    color: #fff;
                    padding: 12px;
                }
                .select2-results__option {
                    padding: 14px 20px;
                    font-size: 16px;
                    color: #fff;
                    cursor: pointer;
                    transition: background-color 0.2s, color 0.2s;
                    /* Perbaikan untuk warna abu-abu */
                    background-color: rgba(255, 255, 255, 0.1); /* Background default */
                }
                .select2-results__option--highlighted {
                    background-color: #8EC5FC !important;
                    color: #1c1c1c !important;
                    border-radius: 15px;
                    margin: 5px;
                }
                .select2-container--default .select2-selection--single .select2-selection__clear {
                    display: none !important;
                }
                .select2-results__options {
                    overflow-y: auto !important;
                    overflow-x: hidden !important;
                    max-height: 200px;
                }
                button.main-button {
                    padding: 15px 30px;
                    border-radius: 18px;
                    border: none;
                    background: linear-gradient(180deg, #8EC5FC 0%, #E0C3FC 100%);
                    color: #1c1c1c;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.3s ease-in-out;
                    box-shadow: 0 5px 20px rgba(142, 197,252, 0.3);
                    font-size: 18px;
                    letter-spacing: -0.2px;
                }
                button.main-button:hover {
                    background: linear-gradient(180deg, #E0C3FC 0%, #8EC5FC 100%);
                    transform: translateY(-3px);
                    box-shadow: 0 8px 25px rgba(142, 197, 252, 0.4);
                }
                button.main-button:disabled {
                    background: rgba(255, 255, 255, 0.05);
                    color: rgba(255, 255, 255, 0.4);
                    border: 1px solid rgba(255,255,255,0.1);
                    cursor: not-allowed;
                    box-shadow: none;
                    transform: translateY(0);
                }
                .footer {
                   flex-shrink: 0; /* Mencegah footer menyusut */
                    padding: 20px 0;
                    text-align: center;
                    color: rgba(255, 255, 255, 0.6);
                    font-size: 14px;
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(5px);
                    border-top: 1px solid rgba(255, 255, 255, 0.1);
                    position: relative; /* Ubah dari fixed ke relative */
                }
            </style>
        </head>
        <body>
            <header class="header">
                <h1 class="header-title">Shortcut Persuratan Pidsus Buleleng</h1>
                <div class="nav-buttons">
                    <a href="/" class="button-link">Beranda</a>
                    <a href="#" class="button-link" onclick="window.history.back(); return false;">Kembali</a>
                </div>
            </header>
            <main>
                <div class="container">
                    <h2>Pilih Template Surat</h2>
                    <form method="GET" action="/isi">
                        <input type="hidden" name="stage" value="${stage}">
                        <div class="form-group">
                            <select id="templateSelect" name="template" required>
                                <option value="">Cari dan pilih template...</option>
                                ${Object.entries(templates).map(([key, t]) => `
                                    <option value="${key}">${t.label}</option>
                                `).join('')}
                            </select>
                            <button type="submit" id="lanjutButton" class="main-button" disabled>Lanjut</button>
                        </div>
                    </form>
                </div>
            </main>
            <div class="footer">
                <p>Hak Cipta © parwatabayu. Pidsus 2025.</p>
            </div>
            <script>
                $(document).ready(function() {
                    $('#templateSelect').select2({
                        placeholder: "Cari dan pilih template...",
                        allowClear: false,
                        minimumResultsForSearch: 0,
                        dropdownParent: $('body')
                    });
                    $('#templateSelect').on('change', function() {
                        if ($(this).val()) {
                            $('#lanjutButton').prop('disabled', false);
                        } else {
                            $('#lanjutButton').prop('disabled', true);
                        }
                    });
                });
            </script>
        </body>
        </html>
    `);
});

// Halaman: Isi Data sesuai template
app.get('/isi', (req, res) => {
    const stage = req.query.stage;
    const templateKey = req.query.template;
    const config = configs[stage];

    if (!config || !config.templates[templateKey]) {
        return res.redirect('/');
    }

    const t = config.templates[templateKey];

    res.send(`
        <html>
        <head>
            <title>Isi Data - ${t.label}</title>
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <style>
                html, body {
                    overflow: hidden;
                    height: 100%;
                }
                body {
                    background-image: url('https://raw.githubusercontent.com/pidsusbuleleng/pidsusbuleleng/refs/heads/main/Asset/back.jpg');
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed;
                    background-color: #2c2c2c;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    display: flex;
                    flex-direction: column;
                    min-height: 100vh;
                    color: #fff;
                }
                .header {
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    width: 100%;
                    height: 64px; 
                    padding: 20px;
                    background: rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(15px);
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    z-index: 1000;
                    border-bottom: 1px solid rgba(255, 255, 255, 0.2);
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    box-sizing: border-box; /* Memastikan padding termasuk dalam tinggi */
                    
                }
                .header-title {
                    font-size: 24px;
                    font-weight: 700;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                    margin: 0;
                    line-height: 1; /* Menyamakan line-height */
                }
                .nav-buttons {
                    display: flex;
                    gap: 10px;
                }
                .button-link {
                    padding: 12px 20px;
                    border-radius: 12px;
                    border: none;
                    background: rgba(255, 255, 255, 0.15);
                    color: #fff;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: background 0.3s, transform 0.3s;
                    backdrop-filter: blur(10px);
                    border: 1px solid rgba(255,255,255,0.3);
                    text-decoration: none;
                    display: inline-block;
                    text-align: center;
                }
                .button-link:hover {
                    background: rgba(255, 255, 255, 0.25);
                    transform: translateY(-2px);
                }
                main {
                    flex: 1;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    margin-top: 64px;       /* sama dengan tinggi header */
                    height: calc(100vh - 64px); /* sisa tinggi layar */ 
                    padding: 80px;
                    box-sizing: border-box;
                    position: relative;
                    z-index: 1;
                    overflow-y: auto;
                }
                main::before {
                    content: '';
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: linear-gradient(135deg, rgba(0, 0, 0, 0.3) 0%, rgba(0, 0, 0, 0.2) 100%);
                    z-index: -1;
                }
                .container {
                    max-width: 1700px;
                    width: 100%;
                    background: rgba(255, 255, 255, 0.1);
                    border-radius: 25px;
                    box-shadow: 0 10px 40px rgba(0,0,0,0.1), 0 0 0 1px rgba(255,255,255,0.2);
                    padding: 20px;
                    box-sizing: border-box;
                    backdrop-filter: blur(35px);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                    margin-top: 0px;
            
                }
                    
                h2 {
                    text-align: left;
                    color: #fff;
                    font-weight: 700;
                    margin-bottom: 25px;
                    border-bottom: 1px solid rgba(255,255,255,0.3);
                    padding-bottom: 15px;
                    font-size: 32px;
                    letter-spacing: -1px;
                    text-shadow: 0 1px 3px rgba(0,0,0,0.5);
                }
                .form-grid {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
                    gap: 30px;
                }
                
                .form-item { display: flex; flex-direction: column; }
                label {
                    margin-bottom: 10px;
                    color: #fff;
                    font-size: 16px;
                    font-weight: 500;
                    letter-spacing: -0.1px;
                }
                input[type="text"], input[type="number"], input[type="date"] {
                    padding: 16px 18px;
                    border: 1px solid rgba(255, 255, 255, 0.4);
                    border-radius: 15px;
                    font-size: 16px;
                    background: rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(10px);
                    transition: all 0.2s;
                    color: #fff;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
                }
                input[type="text"]:focus, input[type="number"]:focus, input[type="date"]:focus {
                    border: 1.5px solid #8EC5FC;
                    outline: none;
                    box-shadow: 0 0 0 4px rgba(142, 197, 252, 0.2);
                }
                .button-group {
                    display: flex;
                    justify-content: flex-end;
                    gap: 20px;
                    margin-top: 50px;
                }
                button.main-button {
                    padding: 15px 30px;
                    border: none;
                    border-radius: 18px;
                    font-size: 18px;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.3s ease-in-out;
                    letter-spacing: -0.2px;
                    background: linear-gradient(180deg, #8EC5FC 0%, #E0C3FC 100%);
                    color: #1c1c1c;
                    box-shadow: 0 5px 20px rgba(142, 197, 252, 0.3);
                }
                button.main-button:hover {
                    background: linear-gradient(180deg, #E0C3FC 0%, #8EC5FC 100%);
                    transform: translateY(-3px);
                    box-shadow: 0 8px 25px rgba(142, 197, 252, 0.4);
                }
                button.clear {
                    padding: 15px 30px; /* Menyamakan padding */
                    border-radius: 18px; /* Menyamakan rounded dengan tombol Generate */
                    border: none;
                    font-size: 18px;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.3s ease-in-out;
                    letter-spacing: -0.2px;
                    background: rgba(255, 255, 255, 0.1);
                    color: #fff;
                    border: 1px solid rgba(255,255,255,0.5);
                    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                    backdrop-filter: blur(10px);
                }
                button.clear:hover {
                    background: rgba(255, 255, 255, 0.2);
                    transform: translateY(-3px);
                    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
                }

                /* Tambahan: tata letak baris saksi dan gaya tombol Hapus agar mirip tombol lain */
                .saksi-row {
                    /* jangan ubah layout .saksi-row (biar tetap mengikuti grid/form-grid) */
                }
                .remove-saksi {
                    display: inline-block;
                    padding: 10px 14px;
                    margin-top: 18px; /* sejajarkan dengan input */
                    margin-left: 8px;
                    border-radius: 14px;
                    border: none;
                    font-size: 14px;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.18s ease-in-out;
                    background: rgba(255, 255, 255, 0.12);
                    color: #fff;
                    border: 1px solid rgba(255,255,255,0.45);
                    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
                    backdrop-filter: blur(8px);
                }
                .remove-saksi:hover {
                    background: rgba(255, 255, 255, 0.2);
                    transform: translateY(-2px);
                    box-shadow: 0 4px 12px rgba(0,0,0,0.06);
                }
                .footer {
                    flex-shrink: 0; /* Mencegah footer menyusut */
                    padding: 20px 0;
                    text-align: center;
                    color: rgba(255, 255, 255, 0.6);
                    font-size: 14px;
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(5px);
                    border-top: 1px solid rgba(255, 255, 255, 0.1);
                    position: relative; /* Ubah dari fixed ke relative */
                }
            </style>
            <script>
                function clearForm() {
                    document.querySelector('form').reset();
                }

                document.addEventListener('DOMContentLoaded', function() {
                    var hasDynamicTable = ${t.hasDynamicTable ? 'true' : 'false'};
                    if (hasDynamicTable) {
                        document.getElementById('saksi-section').style.display = 'block';
                    } else {
                        document.getElementById('saksi-section').style.display = 'none';
                    }

                    var addSaksiBtn = document.getElementById('add-saksi');
                    var dynamicSaksiFields = document.getElementById('dynamic-saksi-fields');

                    if (addSaksiBtn && dynamicSaksiFields) {
                        addSaksiBtn.addEventListener('click', function() {
                            var newSaksiRow = document.createElement('div');
                            newSaksiRow.className = 'saksi-row';
                            newSaksiRow.innerHTML = ''
                                + '<div class="form-item">'
                                + '<label>Nama Saksi:</label>'
                                + '<input type="text" name="namasaksi[]" required />'
                                + '</div>'
                                + '<div class="form-item">'
                                + '<label>Alamat/Jabatan:</label>'
                                + '<input type="text" name="alamatjabatan[]" required />'
                                + '</div>'
                                + '<button type="button" class="remove-saksi">Hapus</button>';
                            dynamicSaksiFields.appendChild(newSaksiRow);
                        });

                        dynamicSaksiFields.addEventListener('click', function(e) {
                            if (e.target && e.target.classList.contains('remove-saksi')) {
                                if (document.querySelectorAll('.saksi-row').length > 1) {
                                    e.target.closest('.saksi-row').remove();
                                } else {
                                    alert("Minimal harus ada satu saksi.");
                                }
                            }
                        });
                    }
                    const tglInput = document.querySelector("input[name='tglsidang']");
                    const hariInput = document.querySelector("input[name='harisidang']");

                    if (tglInput && hariInput) {
                        tglInput.addEventListener("change", function () {
                            const date = new Date(this.value);
                            const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

                            if (!isNaN(date)) {
                                hariInput.value = days[date.getDay()];
                            } else {
                                hariInput.value = "";
                            }
                        });
                    }
                });
               
            </script>
        </head>
        <body>
            <header class="header">
                <h1 class="header-title">Shortcut Persuratan Pidsus Buleleng</h1>
                <div class="nav-buttons">
                    <a href="/" class="button-link">Beranda</a>
                    <a href="#" class="button-link" onclick="window.history.back(); return false;">Kembali</a>
                </div>
            </header>
            <main>
                <div class="container">
                    <h2>Isi Data - ${t.label}</h2>
                    <form method="POST" action="/generate">
                        <input type="hidden" name="stage" value="${stage}" />
                        <input type="hidden" name="template" value="${templateKey}" />
                        
                        <div class="form-grid">
                            ${t.fields.map(f => `
                                <div class="form-item">
                                    <label>${f.label}:</label>
                                    <input name="${f.name}" type="${f.type}" ${f.min !== undefined ? `min="${f.min}"` : ''} required />
                                </div>
                            `).join('')}
                        </div>

                        <div id="saksi-section" style="display:none;">
                        <hr style="border: 1px solid rgba(255,255,255,0.2); margin: 10px 0;">
                        <h2>Daftar Saksi</h2>
                        <div id="dynamic-saksi-fields" class="form-grid">
                            <div class="saksi-row">
                                <div class="form-item">
                                    <label>Nama Saksi:</label>
                                    <input type="text" name="namasaksi[]" />
                                </div>
                                <div class="form-item">
                                    <label>Alamat/Jabatan:</label>
                                    <input type="text" name="alamatjabatan[]" />
                                </div>
                                <button type="button" class="remove-saksi">Hapus</button>
                            </div>
                        </div>
                        <button type="button" id="add-saksi" class="main-button" style="margin-top: 20px;">Tambah Saksi</button>
                    </div>

                        <div class="button-group">
                            <button type="button" class="clear" onclick="clearForm()">Clear</button>
                            <button type="submit" class="main-button">Generate</button>
                        </div>
                    </form>
                </div>
            </main>
            <div class="footer">
                <p>Hak Cipta © parwatabayu. Pidsus 2025.</p>
            </div>
        </body>
        </html>
    `);
});

app.post('/generate', (req, res) => {
    const stage = req.body.stage;
    const templateKey = req.body.template;
    const config = configs[stage];

    if (!config || !config.templates[templateKey]) {
        return res.status(400).send('Template tidak ditemukan.');
    }

    // Perbaikan: Mendefinisikan variabel 't' dari konfigurasi
    const t = config.templates[templateKey];
    const templatePath = path.join(__dirname, 'templates', stage, t.file);

    let content;
    try {
        content = fs.readFileSync(templatePath);
    } catch (err) {
        return res.status(500).send(`Template file tidak ditemukan. Pastikan file ada di folder /templates/${stage}/${t.file}.`);
    }

    const data = {};
    t.fields.forEach(f => {
        // Jika field tipe date, format ke bahasa Indonesia
        if (f.type === 'date') {
            data[f.name] = formatDateID(req.body[f.name]);
        } else {
            data[f.name] = req.body[f.name] !== undefined ? req.body[f.name] : '';
        }
    });

    // Jika ada field namasaksi sebagai input tunggal, hilangkan koma trailing dan rapikan spasi
    if (data.namasaksi) {
        data.namasaksi = String(data.namasaksi).trim().replace(/,+$/g, '').replace(/\s+/g, ' ');
    }
    
    // Siapkan variabel saksiArr pada scope yang lebih luas (dipakai juga saat fallback)
    let saksiArr = null;
    // Jika template mempunyai tabel dinamis (saksi), susun array untuk Docxtemplater
    if (t.hasDynamicTable) {
         // dukung berbagai variasi nama input: namasaksi / nama_saksi, alamatjabatan / alamat_jabatan
         const rawNama = req.body.namasaksi || req.body['namasaksi'];
         const rawAlamat = req.body.alamatjabatan || req.body['alamat_jabatan'];
 
         const arrNama = Array.isArray(rawNama) ? rawNama : (rawNama ? [rawNama] : []);
         const arrAlamat = Array.isArray(rawAlamat) ? rawAlamat : (rawAlamat ? [rawAlamat] : []);
 
         // sanitize name values: trim, remove trailing commas, collapse spaces
         const normalizeName = (arr) =>
             arr.map(v => {
                 let s = (typeof v === 'string' ? v.trim() : (v == null ? '' : String(v)));
                 s = s.replace(/,+$/g, ''); // remove trailing commas
                 s = s.replace(/\s+/g, ' '); // collapse spaces
                 return s;
             });
         // generic normalize for other fields
         const normalize = (arr) =>
             arr.map(v => (typeof v === 'string' ? v.trim() : (v == null ? '' : String(v))));
 
         const tNama = normalizeName(arrNama);
         const tAlamat = normalize(arrAlamat);
 
         // build rows, skip rows where both name and address are empty
         const tmp = [];
         const len = Math.max(tNama.length, tAlamat.length);
         for (let i = 0; i < len; i++) {
             const name = tNama[i] || '';
             const addr = tAlamat[i] || '';
             if (name === '' && addr === '') continue; // drop fully empty rows
             tmp.push({ namasaksi: name, alamatjabatan: addr });
         }
 
         // If any entries exist, reindex them and add 'no' property to avoid Word auto-number gaps
         saksiArr = tmp.map((item, idx) => ({
            no: idx + 1,
            namasaksi: item.namasaksi || '',
            alamatjabatan: item.alamatjabatan || ''
         }));
 
         // assign to data for Docxtemplater
         data.saksi = saksiArr;
 
         // also set single placeholders (first entry) for templates expecting single values
         data.namasaksi = saksiArr[0] ? saksiArr[0].namasaksi : '';
         data.alamatjabatan = saksiArr[0] ? saksiArr[0].alamatjabatan : '';
     }
 
    let doc;
    try {
        const zip = new PizZip(content);
        doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.setData(data);
        doc.render();
    } catch (error) {
        // Jika template menggunakan raw tag untuk 'saksi' (mengharapkan string), konversi array jadi string lalu coba ulang
        if (error && error.properties && error.properties.id === 'invalid_raw_xml_value' && error.properties.xtag === 'saksi' && saksiArr) {
            // escape XML special chars
            const escapeXml = (str) => String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
 
            // Buat representasi teks sederhana: "1. Name - Address" per baris
            const rowsText = saksiArr.map(r => `${r.no}. ${escapeXml(r.namasaksi)} - ${escapeXml(r.alamatjabatan)}`).join('\n');
            data.saksi = rowsText;
 
            // Coba render ulang dengan data.saksi sebagai string
            try {
                const zip2 = new PizZip(content);
                doc = new Docxtemplater(zip2, { paragraphLoop: true, linebreaks: true });
                doc.setData(data);
                doc.render();
            } catch (err2) {
                return res.status(500).send('Template error (after fallback): ' + err2.message);
            }
        } else {
            return res.status(500).send('Template error: ' + error.message);
        }
    }

    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    // Ambil data tanggal dari request body
    const tglSidangRaw = req.body.tglsidang;

    // Format tanggal ke DD-MM-YYYY
    let formattedDate = '';
    if (tglSidangRaw) {
        const [year, month, day] = tglSidangRaw.split('-');
        formattedDate = `${day}-${month}-${year}`;
    }

    // sanitize helper for filenames (remove chars not allowed in filenames)
    const sanitizeForFilename = (str) => {
        if (!str) return '';
        return String(str)
            // Hapus karakter ilegal untuk filename dan koma
            .replace(/[\/\\:*?"<>|,]+/g, '')
            // gabungkan spasi berlebih jadi satu
            .replace(/\s+/g, ' ')
            .trim();
    };

    // Gabungkan nama template dan tanggal untuk nama file
    let filename;
    if (templateKey === 'P-37 SAKSI BRI') {
        // prefer first entry from saksiArr (built earlier), then data.namasaksi, then raw request
        const firstSaksi = (saksiArr && saksiArr[0] && saksiArr[0].namasaksi)
            ? saksiArr[0].namasaksi
            : (data.namasaksi || req.body.namasaksi || '');
        const namePart = sanitizeForFilename(firstSaksi);
        if (namePart) {
            filename = `${templateKey} ${namePart} ${formattedDate}.docx`;
        } else {
            filename = `${templateKey} ${formattedDate}.docx`;
        }
    } else {
        filename = `${templateKey} ${formattedDate}.docx`;
    }

    res.set('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buf);
    });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
});

