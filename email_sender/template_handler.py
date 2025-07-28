"""
Template HTML untuk email disposisi
"""
import os
from pathlib import Path
from jinja2 import Template

# Get the path to the logo image if exists
CURRENT_DIR = Path(__file__).parent.parent
LOGO_PATH = CURRENT_DIR / "kop.jpeg"

def get_email_template():
    """
    Returns the base HTML template for emails
    """
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333;
                margin: 0;
                padding: 0;
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                text-align: center;
                padding: 20px;
                background-color: #f8f9fa;
                border-bottom: 3px solid #003366;
            }
            .header img {
                max-width: 100%;
                height: auto;
                margin-bottom: 15px;
            }
            .content {
                padding: 30px 20px;
                background-color: #ffffff;
            }
            .meta-info {
                background-color: #f8f9fa;
                padding: 15px;
                border-radius: 5px;
                margin: 20px 0;
            }
            .meta-info p {
                margin: 5px 0;
            }
            .instruksi {
                background-color: #e9ecef;
                padding: 15px;
                border-left: 4px solid #003366;
                margin: 20px 0;
            }
            .footer {
                text-align: center;
                padding: 20px;
                font-size: 12px;
                color: #666;
                border-top: 1px solid #dee2e6;
            }
            .klasifikasi {
                color: #dc3545;
                font-weight: bold;
                margin: 10px 0;
            }
            .button {
                display: inline-block;
                padding: 10px 20px;
                background-color: #003366;
                color: white;
                text-decoration: none;
                border-radius: 5px;
                margin-top: 20px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                {% if logo_exists %}
                <img src="cid:logo" alt="Logo">
                {% endif %}
                <h2>LEMBAR DISPOSISI</h2>
            </div>
            
            <div class="content">
                <p>Dengan hormat,</p>
                
                <div class="meta-info">
                    <p><strong>Nomor Surat:</strong> {{ nomor_surat }}</p>
                    <p><strong>Pengirim:</strong> {{ nama_pengirim }}</p>
                    <p><strong>Perihal:</strong> {{ perihal }}</p>
                    <p><strong>Tanggal:</strong> {{ tanggal }}</p>
                </div>

                {% if klasifikasi %}
                <div class="klasifikasi">
                    {{ klasifikasi|join(" | ") }}
                </div>
                {% endif %}

                <div class="instruksi">
                    <h3>Instruksi:</h3>
                    {% for item in instruksi_list %}
                    <p>• {{ item }}</p>
                    {% endfor %}
                </div>

                <p>Mohon untuk segera ditindaklanjuti sesuai dengan disposisi yang diberikan.</p>
                
                {% if link_dokumen %}
                <a href="{{ link_dokumen }}" class="button">Lihat Dokumen</a>
                {% endif %}
            </div>

            <div class="footer">
                <p>Email ini dikirim secara otomatis oleh Sistem Disposisi</p>
                <p>© {{ tahun }} Sistem Disposisi JCC</p>
            </div>
        </div>
    </body>
    </html>
    """

def render_email_template(data):
    """
    Render email template with provided data
    
    Args:
        data (dict): Dictionary containing:
            - nomor_surat (str): Nomor surat
            - nama_pengirim (str): Nama pengirim
            - perihal (str): Perihal surat
            - tanggal (str): Tanggal surat
            - klasifikasi (list): List klasifikasi surat [RAHASIA, PENTING, SEGERA]
            - instruksi_list (list): List instruksi disposisi
            - link_dokumen (str, optional): Link ke dokumen
            
    Returns:
        str: Rendered HTML email content
    """
    template = Template(get_email_template())
    
    # Add default values and format data
    data.setdefault('tahun', '2025')
    data.setdefault('logo_exists', LOGO_PATH.exists())
    data.setdefault('link_dokumen', '')
    
    # Ensure lists are properly formatted
    data.setdefault('klasifikasi', [])
    if isinstance(data.get('instruksi'), str):
        data['instruksi_list'] = [x.strip() for x in data['instruksi'].split('\n') if x.strip()]
    else:
        data['instruksi_list'] = data.get('instruksi_list', [])
    
    return template.render(**data)
