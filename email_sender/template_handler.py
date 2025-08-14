# email_sender/template_handler.py - ENHANCED PROFESSIONAL VERSION

"""
Enhanced Professional HTML Template untuk email disposisi dengan desain yang lebih clean dan modern
Features:
- Clean, minimalist design
- Professional color scheme  
- Better typography and spacing
- Responsive layout
- Improved accessibility
- Modern card-based layout
"""
import os
from jinja2 import Template
def get_enhanced_professional_template():
    """
    Returns the enhanced professional HTML template for emails
    """
    return """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="format-detection" content="telephone=no">
        <title>Lembar Disposisi - {{ nomor_surat or 'N/A' }}</title>
        <style>
            /* Reset and base styles */
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', Arial, sans-serif;
                line-height: 1.6;
                color: #1a1a1a;
                background-color: #f8fafc;
                margin: 0;
                padding: 0;
                -webkit-text-size-adjust: 100%;
                -ms-text-size-adjust: 100%;
            }
            
            /* Email wrapper */
            .email-wrapper {
                background-color: #f8fafc;
                padding: 20px 10px;
                min-height: 100vh;
            }
            
            /* Main container */
            .container {
                max-width: 680px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05), 0 1px 3px rgba(0, 0, 0, 0.1);
                overflow: hidden;
                border: 1px solid #e5e7eb;
            }
            
            /* Header Section */
            .header {
                background: linear-gradient(135deg, #1f2937 0%, #374151 100%);
                color: white;
                padding: 32px 24px;
                text-align: center;
                position: relative;
            }
            
            .header-content {
                position: relative;
                z-index: 2;
            }
            .header h1 {
                font-size: 28px;
                font-weight: 700;
                margin-bottom: 8px;
                letter-spacing: -0.025em;
            }
            
            .header .subtitle {
                font-size: 14px;
                opacity: 0.9;
                font-weight: 400;
                text-transform: uppercase;
                letter-spacing: 0.05em;
            }
            
            /* Content Section */
            .content {
                padding: 32px 24px;
            }
            
            .greeting {
                font-size: 16px;
                margin-bottom: 24px;
                color: #374151;
                font-weight: 500;
            }
            
            /* Card components */
            .card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin: 20px 0;
                overflow: hidden;
            }
            
            .card-header {
                background: #f9fafb;
                border-bottom: 1px solid #e5e7eb;
                padding: 16px 20px;
            }
            
            .card-header h3 {
                font-size: 16px;
                font-weight: 600;
                color: #1f2937;
                margin: 0;
                display: flex;
                align-items: center;
            }
            
            .card-header .icon {
                margin-right: 8px;
                font-size: 16px;
            }
            
            .card-body {
                padding: 20px;
            }
            
            /* Document Info Grid */
            .info-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 16px;
                margin: 0;
            }
            
            .info-item {
                display: flex;
                flex-direction: column;
            }
            
            .info-label {
                font-size: 11px;
                font-weight: 600;
                color: #6b7280;
                text-transform: uppercase;
                letter-spacing: 0.05em;
                margin-bottom: 4px;
            }
            
            .info-value {
                font-size: 14px;
                color: #1f2937;
                font-weight: 500;
                background-color: #f9fafb;
                padding: 8px 12px;
                border-radius: 4px;
                border: 1px solid #e5e7eb;
                min-height: 20px;
            }
            
            /* Classification and Status Tags */
            .tags-container {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                margin: 16px 0;
            }
            
            .tag {
                display: inline-flex;
                align-items: center;
                padding: 4px 12px;
                border-radius: 12px;
                font-size: 11px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.05em;
            }
            
            .tag-urgent {
                background: #fee2e2;
                color: #991b1b;
                border: 1px solid #fecaca;
            }
            
            .tag-important {
                background: #fef3c7;
                color: #92400e;
                border: 1px solid #fde68a;
            }
            
            .tag-confidential {
                background: #ddd6fe;
                color: #5b21b6;
                border: 1px solid #c4b5fd;
            }
            
            /* Recipients Section */
            .recipients-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
                gap: 8px;
                margin: 16px 0;
            }
            
            .recipient-item {
                background: #f0f9ff;
                border: 1px solid #bae6fd;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 13px;
                font-weight: 500;
                color: #0c4a6e;
                text-align: center;
            }
            
            /* Instructions Section */
            .instructions-list {
                list-style: none;
                margin: 16px 0;
                padding: 0;
            }
            
            .instruction-item {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                padding: 12px 16px;
                margin-bottom: 8px;
                font-size: 14px;
                color: #475569;
                position: relative;
                padding-left: 40px;
            }
            
            .instruction-item::before {
                content: "•";
                position: absolute;
                left: 16px;
                top: 12px;
                color: #3b82f6;
                font-weight: bold;
                font-size: 16px;
            }
            
            /* Action Section */
            .action-section {
                background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                color: white;
                padding: 20px;
                border-radius: 6px;
                margin: 24px 0;
                text-align: center;
            }
            
            .action-section .icon {
                font-size: 24px;
                margin-bottom: 8px;
                display: block;
            }
            
            .action-section p {
                font-size: 15px;
                font-weight: 500;
                margin: 0;
            }
            
            /* Footer */
            .footer {
                background-color: #f9fafb;
                border-top: 1px solid #e5e7eb;
                color: #6b7280;
                text-align: center;
                padding: 24px 20px;
                font-size: 12px;
            }
            
            .footer-content {
                max-width: 500px;
                margin: 0 auto;
            }
            
            .footer-logo {
                font-weight: 600;
                color: #374151;
                margin-bottom: 8px;
            }
            
            .footer-text {
                line-height: 1.5;
            }
            
            .footer-divider {
                width: 40px;
                height: 1px;
                background: #d1d5db;
                margin: 12px auto;
            }
            
            /* Responsive Design */
            @media only screen and (max-width: 600px) {
                .email-wrapper {
                    padding: 10px 5px;
                }
                
                .container {
                    border-radius: 0;
                    margin: 0;
                    box-shadow: none;
                }
                
                .header {
                    padding: 24px 16px;
                }
                
                .header h1 {
                    font-size: 24px;
                }
                
                .content {
                    padding: 24px 16px;
                }
                
                .info-grid {
                    grid-template-columns: 1fr;
                }
                
                .recipients-grid {
                    grid-template-columns: 1fr;
                }
                
                .card-body {
                    padding: 16px;
                }
                
                .footer {
                    padding: 20px 16px;
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
                    border: none;
                    border-radius: 0;
                }
                
                .header {
                    background: #1f2937 !important;
                    -webkit-print-color-adjust: exact;
                    color-adjust: exact;
                }
                
                .card {
                    break-inside: avoid;
                }
            }
            
            /* Dark mode support */
            @media (prefers-color-scheme: dark) {
                .info-value {
                    background-color: #f3f4f6;
                }
                
                .instruction-item {
                    background: #f1f5f9;
                }
            }
        </style>
    </head>
    <body>
        <div class="email-wrapper">
            <div class="container">
                <!-- Header Section -->
                <div class="header">
                    <div class="header-content">
                        <h1>LEMBAR DISPOSISI</h1>
                        <div class="subtitle">Sistem Manajemen Surat Digital</div>
                    </div>
                </div>
                
                <!-- Content Section -->
                <div class="content">
                    <div class="greeting">
                        <strong>Kepada Yth. Tim Terkait,</strong>
                    </div>
                    
                    <!-- Document Information Card -->
                    <div class="card">
                        <div class="card-header">
                            <h3><span class="icon">📄</span>Informasi Dokumen</h3>
                        </div>
                        <div class="card-body">
                            <div class="info-grid">
                                <div class="info-item">
                                    <div class="info-label">Nomor Surat</div>
                                    <div class="info-value">{{ nomor_surat or '-' }}</div>
                                </div>
                                <div class="info-item">
                                    <div class="info-label">Tanggal</div>
                                    <div class="info-value">{{ tanggal or '-' }}</div>
                                </div>
                                <div class="info-item">
                                    <div class="info-label">Pengirim</div>
                                    <div class="info-value">{{ nama_pengirim or '-' }}</div>
                                </div>
                                <div class="info-item">
                                    <div class="info-label">Perihal</div>
                                    <div class="info-value">{{ perihal or '-' }}</div>
                                </div>
                            </div>
                            
                            <!-- Classification Tags -->
                            {% if klasifikasi and klasifikasi|length > 0 %}
                            <div class="tags-container">
                                {% for item in klasifikasi %}
                                <span class="tag {% if item == 'SEGERA' %}tag-urgent{% elif item == 'PENTING' %}tag-important{% elif item == 'RAHASIA' %}tag-confidential{% endif %}">
                                    {{ item }}
                                </span>
                                {% endfor %}
                            </div>
                            {% endif %}
                        </div>
                    </div>

                    <!-- Recipients Card -->
                    {% if disposisi_kepada and disposisi_kepada|length > 0 %}
                    <div class="card">
                        <div class="card-header">
                            <h3><span class="icon">👥</span>Disposisi Kepada</h3>
                        </div>
                        <div class="card-body">
                            <div class="recipients-grid">
                                {% for posisi in disposisi_kepada %}
                                <div class="recipient-item">{{ posisi }}</div>
                                {% endfor %}
                            </div>
                        </div>
                    </div>
                    {% endif %}

                    <!-- Selected Recipients (for enhanced template) -->
                    {% if selected_recipients and selected_recipients|length > 0 %}
                    <div class="card">
                        <div class="card-header">
                            <h3><span class="icon">📧</span>Penerima Email</h3>
                        </div>
                        <div class="card-body">
                            <div class="recipients-grid">
                                {% for recipient in selected_recipients %}
                                <div class="recipient-item">{{ recipient }}</div>
                                {% endfor %}
                            </div>
                        </div>
                    </div>
                    {% endif %}

                    <!-- Instructions Card -->
                    {% if instruksi_list and instruksi_list|length > 0 %}
                    <div class="card">
                        <div class="card-header">
                            <h3><span class="icon">📋</span>Instruksi Disposisi</h3>
                        </div>
                        <div class="card-body">
                            <ul class="instructions-list">
                                {% for item in instruksi_list %}
                                <li class="instruction-item">{{ item }}</li>
                                {% endfor %}
                            </ul>
                        </div>
                    </div>
                    {% endif %}

                    <!-- Action Section -->
                    <div class="action-section">
                        <span class="icon">⚡</span>
                        <p>Mohon untuk segera ditindaklanjuti sesuai dengan disposisi yang diberikan.</p>
                    </div>
                    
                </div>

                <!-- Footer -->
                <div class="footer">
                    <div class="footer-content">
                        <div class="footer-logo">
                            🏢 Sistem Disposisi Digital
                        </div>
                        <div class="footer-divider"></div>
                        <div class="footer-text">
                            Email ini dikirim secara otomatis oleh sistem<br>
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
    Enhanced notification template with clean design
    """
    return """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ title or 'Notifikasi Sistem' }}</title>
        <style>
            body {
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
                line-height: 1.6;
                color: #1f2937;
                background-color: #f8fafc;
                margin: 0;
                padding: 20px;
            }
            
            .container {
                max-width: 480px;
                margin: 0 auto;
                background-color: #ffffff;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
                overflow: hidden;
                border: 1px solid #e5e7eb;
            }
            
            .header {
                background: linear-gradient(135deg, #3b82f6, #1d4ed8);
                color: white;
                text-align: center;
                padding: 24px 20px;
            }
            
            .header h2 {
                margin: 0;
                font-size: 20px;
                font-weight: 600;
            }
            
            .content {
                padding: 24px;
                text-align: center;
            }
            
            .icon {
                font-size: 48px;
                margin-bottom: 16px;
                display: block;
            }
            
            .message {
                font-size: 15px;
                color: #374151;
                margin: 0;
                line-height: 1.6;
            }
            
            .footer {
                background-color: #f9fafb;
                padding: 16px;
                text-align: center;
                font-size: 12px;
                color: #6b7280;
                border-top: 1px solid #e5e7eb;
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
    Render enhanced email template with provided data
    """
    # Use enhanced template by default
    template = Template(get_enhanced_professional_template())
    
    # Add default values and format data
    data.setdefault('tahun', '2025')
    
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

# Enhanced notification functions
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

# Enhanced template features
def get_template_features():
    """
    Returns information about the enhanced template features
    """
    return {
        'design': 'Clean and professional with modern card-based layout',
        'colors': 'Professional color scheme with proper contrast',
        'typography': 'Enhanced typography with Segoe UI font stack',
        'responsive': 'Fully responsive design for all devices',
        'accessibility': 'WCAG compliant with proper contrast ratios',
        'features': [
            'Card-based information layout',
            'Clean classification tags',
            'Professional recipient display',
            'Enhanced instruction formatting',
            'Responsive grid system',
            'Print-friendly styles',
            'Dark mode considerations'
        ]
    }

# Test function for the enhanced template
def preview_enhanced_template():
    """Generate preview HTML for testing the enhanced template"""
    sample_data = {
        'nomor_surat': 'DS/001/2025',
        'nama_pengirim': 'PT. Contoh Perusahaan',
        'perihal': 'Laporan Keuangan Triwulan IV Tahun 2024',
        'tanggal': '15 Januari 2025',
        'klasifikasi': ['PENTING', 'SEGERA'],
        'disposisi_kepada': [
            'Direktur Utama',
            'Direktur Keuangan', 
            'Manager ops, adm'
        ],
        'selected_recipients': [
            'Direktur Utama',
            'Direktur Keuangan',
            'Manager ops',
            'Manager adm'
        ],
        'instruksi_list': [
            'Mohon dipelajari dan berikan tanggapan dalam 3 hari kerja',
            'Koordinasikan dengan tim terkait untuk analisis mendalam',
            'Siapkan presentasi untuk rapat direksi minggu depan',
            'CC hasil review ke Direktur Utama'
        ]
    }
    
    html_content = render_email_template(sample_data)
    
    # Save preview file
    preview_path = os.path.join(os.path.dirname(__file__), "enhanced_email_preview.html")
    with open(preview_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Enhanced template preview saved: {preview_path}")
    return html_content

if __name__ == "__main__":
    preview_enhanced_template()
    features = get_template_features()
    print("\nEnhanced Template Features:")
    for key, value in features.items():
        if isinstance(value, list):
            print(f"{key.title()}:")
            for item in value:
                print(f"  • {item}")
        else:
            print(f"{key.title()}: {value}")