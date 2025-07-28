def get_email_template(position, nama_pengirim, nomor_surat, perihal, instruksi):
    """
    Menghasilkan template email HTML profesional untuk notifikasi disposisi
    dengan branding PT Jasamarga Jalanlayang Cikampek.
    """
    html_content = f"""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Notifikasi Disposisi Surat</title>
        <style>
            body {{
                font-family: 'Segoe UI', Calibri, Arial, sans-serif;
                line-height: 1.6;
                color: #333333;
                background-color: #f4f7f6;
                margin: 0;
                padding: 0;
            }}
            .container {{
                max-width: 650px;
                margin: 20px auto;
                background-color: #ffffff;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            }}
            .header {{
                background-color: #00529B; /* Warna biru korporat Jasa Marga */
                color: white;
                padding: 25px;
                text-align: center;
            }}
            .header img.logo {{
                max-width: 180px;
                margin-bottom: 15px;
            }}
            .header h2 {{
                margin: 0;
                font-size: 24px;
            }}
            .content {{
                padding: 30px;
            }}
            .content p {{
                margin-bottom: 15px;
            }}
            .details-table {{
                width: 100%;
                border-collapse: collapse;
                margin: 25px 0;
            }}
            .details-table td {{
                padding: 12px;
                border-bottom: 1px solid #eeeeee;
            }}
            .details-table td.label {{
                font-weight: bold;
                width: 30%;
                color: #555555;
            }}
            .instruction-box {{
                background-color: #e7f5ff; /* Latar belakang biru muda untuk instruksi */
                border-left: 5px solid #00529B; /* Aksen biru di kiri */
                padding: 20px;
                margin: 25px 0;
                border-radius: 5px;
            }}
            .instruction-box h3 {{
                margin-top: 0;
                color: #00529B;
            }}
            .footer {{
                text-align: center;
                padding: 20px;
                font-size: 12px;
                color: #888888;
                background-color: #f4f7f6;
            }}
            .footer .brand {{
                font-weight: bold;
                color: #555555;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="https://www.jasamarga.com/assets/img/logo/logo-jasamarga.png" alt="Logo Jasa Marga" class="logo">
                <h2>Notifikasi Disposisi Surat</h2>
            </div>
            <div class="content">
                <p>Yth. <strong>{position}</strong>,</p>
                <p>Anda telah menerima disposisi untuk surat berikut yang diteruskan oleh <strong>{nama_pengirim}</strong>.</p>
                
                <table class="details-table">
                    <tr>
                        <td class="label">Nomor Surat</td>
                        <td>{nomor_surat}</td>
                    </tr>
                    <tr>
                        <td class="label">Perihal</td>
                        <td>{perihal}</td>
                    </tr>
                </table>
                
                <div class="instruction-box">
                    <h3>Instruksi / Arahan</h3>
                    <p>{instruksi}</p>
                </div>
                
                <p>Mohon untuk dapat segera menindaklanjuti arahan tersebut. Terima kasih.</p>
            </div>
            <div class="footer">
                <p class="brand">PT Jasamarga Jalanlayang Cikampek</p>
                <p>Email ini dikirim secara otomatis melalui Sistem Disposisi Surat.</p>
            </div>
        </div>
    </body>
    </html>
    """
    return html_content