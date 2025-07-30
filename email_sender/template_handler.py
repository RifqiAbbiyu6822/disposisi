# email_sender/template_handler.py

"""
Professional HTML Template untuk email disposisi
"""
import os
from pathlib import Path
from jinja2 import Template

# Get the path to the logo image if it exists
# This path assumes 'kop.jpeg' is in the root of the project, one level above the 'email_sender' directory.
CURRENT_DIR = Path(__file__).parent.parent
LOGO_PATH = CURRENT_DIR / "kop.jpeg"

def get_email_template():
    """
    Returns the professional HTML template for emails
    """
    return """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Lembar Disposisi</title>
        <style>
            /* Reset and base styles */
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #2c3e50;
                background-color: #f8f9fa;
                margin: 0;
                padding: 0;
            }
            
            .email-wrapper {
                background-color: #f8f9fa;
                padding: 20px 10px;
                min-height: 100vh;
            }
            
            .container {
                max-width: 700px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 12px;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
                overflow: hidden;
            }
            
            /* Header Section */
            .header {
                background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
                color: white;
                text-align: center;
                padding: 40px 20px;
                position: relative;
            }
            
            .header::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><pattern id="grid" width="10" height="10" patternUnits="userSpaceOnUse"><path d="M 10 0 L 0 0 0 10" fill="none" stroke="rgba(255,255,255,0.05)" stroke-width="1"/></pattern></defs><rect width="100" height="100" fill="url(%23grid)"/></svg>');
                opacity: 0.3;
            }
            
            .header-content {
                position: relative;
                z-index: 1;
            }
            
            .header img {
                max-width: 200px;
                height: auto;
                margin-bottom: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
            }
            
            .header h1 {
                font-size: 28px;
                font-weight: 700;
                margin-bottom: 8px;
                text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
            }
            
            .header .subtitle {
                font-size: 16px;
                opacity: 0.9;
                font-weight: 300;
            }
            
            /* Content Section */
            .content {
                padding: 40px 30px;
            }
            
            .greeting {
                font-size: 16px;
                margin-bottom: 30px;
                color: #2c3e50;
            }
            
            /* Document Info Card */
            .document-info {
                background: linear-gradient(135deg, #ecf0f1 0%, #f8f9fa 100%);
                border-radius: 12px;
                padding: 25px;
                margin: 25px 0;
                border-left: 5px solid #3498db;
                box-shadow: 0 3px 10px rgba(0, 0, 0, 0.08);
            }
            
            .document-info h3 {
                color: #2c3e50;
                font-size: 18px;
                margin-bottom: 15px;
                display: flex;
                align-items: center;
            }
            
            .document-info h3::before {
                content: "📄";
                margin-right: 10px;
                font-size: 20px;
            }
            
            .info-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 15px;
            }
            
            .info-item {
                display: flex;
                flex-direction: column;
            }
            
            .info-label {
                font-size: 12px;
                font-weight: 600;
                color: #7f8c8d;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 5px;
            }
            
            .info-value {
                font-size: 15px;
                color: #2c3e50;
                font-weight: 500;
                background-color: #ffffff;
                padding: 8px 12px;
                border-radius: 6px;
                border: 1px solid #e9ecef;
            }
            
            /* Classification Tags */
            .classification {
                margin: 20px 0;
            }
            
            .classification-tags {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                margin-top: 10px;
            }
            
            .tag {
                background: linear-gradient(135deg, #e74c3c, #c0392b);
                color: white;
                padding: 6px 14px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                box-shadow: 0 2px 8px rgba(231, 76, 60, 0.3);
            }
            
            /* Instructions Section */
            .instructions {
                background: linear-gradient(135deg, #f39c12 0%, #e67e22 100%);
                color: white;
                border-radius: 12px;
                padding: 25px;
                margin: 25px 0;
                box-shadow: 0 5px 15px rgba(243, 156, 18, 0.3);
            }
            
            .instructions h3 {
                font-size: 20px;
                margin-bottom: 15px;
                display: flex;
                align-items: center;
            }
            
            .instructions h3::before {
                content: "⚡";
                margin-right: 10px;
                font-size: 22px;
            }
            
            .instruction-list {
                list-style: none;
                padding: 0;
            }
            
            .instruction-item {
                background: rgba(255, 255, 255, 0.1);
                margin-bottom: 10px;
                padding: 12px 16px;
                border-radius: 8px;
                border-left: 4px solid rgba(255, 255, 255, 0.3);
                backdrop-filter: blur(10px);
            }
            
            .instruction-item::before {
                content: "▶";
                margin-right: 10px;
                opacity: 0.8;
            }
            
            /* Action Section */
            .action-section {
                background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%);
                color: white;
                padding: 20px;
                border-radius: 12px;
                margin: 25px 0;
                text-align: center;
                box-shadow: 0 5px 15px rgba(39, 174, 96, 0.3);
            }
            
            .action-section p {
                font-size: 16px;
                font-weight: 500;
                margin: 0;
            }
            
            .action-section::before {
                content: "🎯";
                display: block;
                font-size: 30px;
                margin-bottom: 10px;
            }
            
            /* Footer */
            .footer {
                background-color: #34495e;
                color: #ecf0f1;
                text-align: center;
                padding: 30px 20px;
                font-size: 14px;
            }
            
            .footer-content {
                max-width: 600px;
                margin: 0 auto;
            }
            
            .footer-logo {
                margin-bottom: 15px;
            }
            
            .footer-text {
                opacity: 0.8;
                line-height: 1.5;
            }
            
            .footer-divider {
                width: 50px;
                height: 2px;
                background: linear-gradient(90deg, transparent, #3498db, transparent);
                margin: 15px auto;
            }
            
            /* Responsive Design */
            @media (max-width: 600px) {
                .email-wrapper {
                    padding: 10px 5px;
                }
                
                .container {
                    border-radius: 0;
                    margin: 0;
                }
                
                .header {
                    padding: 30px 15px;
                }
                
                .header h1 {
                    font-size: 24px;
                }
                
                .content {
                    padding: 25px 20px;
                }
                
                .info-grid {
                    grid-template-columns: 1fr;
                }
                
                .classification-tags {
                    justify-content: center;
                }
            }
            
            /* Print Styles */
            @media print {
                .email-wrapper {
                    background-color: white;
                    padding: 0;
                }
                
                .container {
                    box-shadow: none;
                    border-radius: 0;
                }
                
                .header {
                    background: #2c3e50 !important;
                    -webkit-print-color-adjust: exact;
                    color-adjust: exact;
                }
            }
        </style>
    </head>
    <body>
        <div class="email-wrapper">
            <div class="container">
                <div class="header">
                    <div class="header-content">
                        {% if logo_exists %}
                        <img src="cid:logo" alt="Logo Perusahaan">
                        {% endif %}
                        <h1>LEMBAR DISPOSISI</h1>
                        <div class="subtitle">Sistem Manajemen Surat Digital</div>
                    </div>
                </div>
                
                <div class="content">
                    <div class="greeting">
                        <strong>Dengan hormat,</strong>
                    </div>
                    
                    <div class="document-info">
                        <h3>Informasi Dokumen</h3>
                        <div class="info-grid">
                            <div class="info-item">
                                <div class="info-label">Nomor Surat</div>
                                <div class="info-value">{{ nomor_surat or '-' }}</div>
                            </div>
                            <div class="info-item">
                                <div class="info-label">Pengirim</div>
                                <div class="info-value">{{ nama_pengirim or '-' }}</div>
                            </div>
                            <div class="info-item">
                                <div class="info-label">Perihal</div>
                                <div class="info-value">{{ perihal or '-' }}</div>
                            </div>
                            <div class="info-item">
                                <div class="info-label">Tanggal</div>
                                <div class="info-value">{{ tanggal or '-' }}</div>
                            </div>
                        </div>
                    </div>

                    {% if klasifikasi and klasifikasi|length > 0 %}
                    <div class="classification">
                        <div class="classification-tags">
                            {% for item in klasifikasi %}
                            <span class="tag">{{ item }}</span>
                            {% endfor %}
                        </div>
                    </div>
                    {% endif %}

                    {% if instruksi_list and instruksi_list|length > 0 %}
                    <div class="instructions">
                        <h3>Instruksi Disposisi</h3>
                        <ul class="instruction-list">
                            {% for item in instruksi_list %}
                            <li class="instruction-item">{{ item }}</li>
                            {% endfor %}
                        </ul>
                    </div>
                    {% endif %}

                    <div class="action-section">
                        <p>Mohon untuk segera ditindaklanjuti sesuai dengan disposisi yang diberikan di atas.</p>
                    </div>
                    
                </div>

                <div class="footer">
                    <div class="footer-content">
                        <div class="footer-logo">
                            🏢 <strong>Sistem Disposisi JCC</strong>
                        </div>
                        <div class="footer-divider"></div>
                        <div class="footer-text">
                            Email ini dikirim secara otomatis oleh Sistem Disposisi Digital<br>
                            © {{ tahun }} - Semua hak dilindungi
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """

def get_notification_template():
    """
    Returns a simple notification template for status updates
    """
    return """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Notifikasi Sistem</title>
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #2c3e50;
                background-color: #f8f9fa;
                margin: 0;
                padding: 20px;
            }
            
            .container {
                max-width: 500px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 12px;
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
                overflow: hidden;
            }
            
            .header {
                background: linear-gradient(135deg, #3498db, #2980b9);
                color: white;
                text-align: center;
                padding: 30px 20px;
            }
            
            .header h2 {
                margin: 0;
                font-size: 24px;
            }
            
            .content {
                padding: 30px;
                text-align: center;
            }
            
            .icon {
                font-size: 48px;
                margin-bottom: 20px;
            }
            
            .message {
                font-size: 16px;
                margin-bottom: 20px;
            }
            
            .footer {
                background-color: #ecf0f1;
                padding: 15px;
                text-align: center;
                font-size: 12px;
                color: #7f8c8d;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>{{ title or 'Notifikasi Sistem' }}</h2>
            </div>
            
            <div class="content">
                <div class="icon">{{ icon or '📧' }}</div>
                <div class="message">{{ message }}</div>
            </div>
            
            <div class="footer">
                Sistem Disposisi Digital © {{ tahun or '2025' }}
            </div>
        </div>
    </body>
    </html>
    """

def render_email_template(data):
    """
    Render email template with provided data
    """
    template = Template(get_email_template())
    
    # Add default values and format data
    data.setdefault('tahun', '2025')
    data.setdefault('logo_exists', LOGO_PATH.exists())
    
    # Ensure lists are properly formatted
    data.setdefault('klasifikasi', [])
    if isinstance(data.get('instruksi'), str):
        data['instruksi_list'] = [x.strip() for x in data['instruksi'].split('\n') if x.strip()]
    else:
        data['instruksi_list'] = data.get('instruksi_list', [])
    
    return template.render(**data)

def render_notification_template(data):
    """
    Render notification template with provided data
    """
    template = Template(get_notification_template())
    
    # Add default values
    data.setdefault('tahun', '2025')
    data.setdefault('title', 'Notifikasi Sistem')
    data.setdefault('icon', '📧')
    
    return template.render(**data)

# Template untuk berbagai jenis notifikasi
def create_success_notification(message):
    """Create a success notification email"""
    return render_notification_template({
        'title': 'Berhasil',
        'icon': '✅',
        'message': message
    })

def create_error_notification(message):
    """Create an error notification email"""
    return render_notification_template({
        'title': 'Terjadi Kesalahan',
        'icon': '❌',
        'message': message
    })

def create_info_notification(message):
    """Create an info notification email"""
    return render_notification_template({
        'title': 'Informasi',
        'icon': 'ℹ️',
        'message': message
    })

# Test function untuk melihat preview template
def preview_template():
    """Generate preview HTML for testing"""
    sample_data = {
        'nomor_surat': 'DS/001/2025',
        'nama_pengirim': 'John Doe',
        'perihal': 'Laporan Keuangan Bulanan',
        'tanggal': '29 Juli 2025',
        'klasifikasi': ['PENTING', 'RAHASIA', 'SEGERA'],
        'instruksi_list': [
            'Mohon dipelajari dan berikan tanggapan',
            'Koordinasikan dengan tim terkait',
            'Laporkan hasil dalam 3 hari kerja',
            'CC kan ke Direktur Utama'
        ]
    }
    
    html_content = render_email_template(sample_data)
    
    # Save preview file
    preview_path = CURRENT_DIR / "email_preview.html"
    with open(preview_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Preview template tersimpan di: {preview_path}")
    return html_content

if __name__ == "__main__":
    preview_template()