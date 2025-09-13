const express = require('express');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');

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
            'surat4_tuntut': {
                file: 'surat4.docx',
                label: 'Surat 4 (Penuntutan)',
                fields: [
                    { name: 'nama', label: 'Nama', type: 'text' },
                    { name: 'tanggal_sidang', label: 'Tanggal Sidang', type: 'text' }
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
                    box-shadow: 0 5px 20px rgba(142, 197, 252, 0.3);
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
                    background-image: url('https://www.iclarified.com/images/news/97556/465557/465557.jpg');
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
                    overflow-y: auto;
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
                    max-width: 1700px;
                    width: 100%;
                    background: rgba(255, 255, 255, 0.1);
                    border-radius: 25px;
                    box-shadow: 0 10px 40px rgba(0,0,0,0.1), 0 0 0 1px rgba(255,255,255,0.2);
                    padding: 40px;
                    box-sizing: border-box;
                    backdrop-filter: blur(35px);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                    margin-top: 60px;
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
                input[type="text"], input[type="number"] {
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
                input[type="text"]:focus, input[type="number"]:focus {
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
    t.fields.forEach(f => data[f.name] = req.body[f.name]);

    let doc;
    try {
        const zip = new PizZip(content);
        doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.setData(data);
        doc.render();
    } catch (error) {
        return res.status(500).send('Template error: ' + error.message);
    }

    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.set('Content-Disposition', `attachment; filename=hasil-${templateKey}.docx`);
    res.send(buf);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
});
