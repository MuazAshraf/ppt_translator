from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import mimetypes
import os
import shutil
from datetime import datetime
import threading
import time
import json
from multi_improved import ALLOWED_FORMATS, translate_pptx_multi


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# Use local output folder instead of temp
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')
app.config['UPLOAD_FOLDER'] = OUTPUT_FOLDER

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Allowed extensions  
ALLOWED_EXTENSIONS = {'pptx'}

# Fallback for translation formats (when multi_improved is available)
ALLOWED_FORMATS = {'pptx', 'pdf'}  # Add more as needed

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/translate', methods=['POST'])
def translate():
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']

        # Check if file is selected
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Only .pptx files are allowed'}), 400

        # Get form data
        service = request.form.get('service', 'google')
        api_key = request.form.get('api_key') or None

        target_langs = request.form.getlist('target_langs')
        if not target_langs:
            fallback_lang = request.form.get('target_lang')
            if fallback_lang:
                target_langs = [fallback_lang]

        target_langs = [lang.strip() for lang in target_langs if lang and lang.strip()]
        if not target_langs:
            return jsonify({'error': 'No target languages provided'}), 400

        # Deduplicate while preserving order
        normalized_langs = []
        seen_langs = set()
        for lang in target_langs:
            if lang not in seen_langs:
                normalized_langs.append(lang)
                seen_langs.add(lang)

        requested_formats = request.form.getlist('formats')
        if not requested_formats:
            requested_formats = ['pptx']

        normalized_formats = []
        for fmt in requested_formats:
            normalized = fmt.lower().strip().lstrip('.')
            if normalized and normalized not in normalized_formats:
                normalized_formats.append(normalized)

        if not normalized_formats:
            normalized_formats.append('pptx')

        # Debug: Print what formats were requested
        print(f"DEBUG: Requested formats from form: {requested_formats}")
        print(f"DEBUG: Normalized formats being used: {normalized_formats}")

        invalid_formats = [fmt for fmt in normalized_formats if fmt not in ALLOWED_FORMATS]
        if invalid_formats:
            return jsonify({'error': f"Invalid output format(s): {', '.join(invalid_formats)}"}), 400

        # Save uploaded file temporarily
        filename = secure_filename(file.filename)
        base_name = filename.rsplit('.', 1)[0]  # Remove extension

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{timestamp}_{filename}")
        output_root = os.path.join(app.config['UPLOAD_FOLDER'], f"{base_name}_{timestamp}")
        zip_download_name = f"{base_name}_{timestamp}.zip"

        file.save(input_path)

        try:
            summary = translate_pptx_multi(
                input_file=input_path,
                target_langs=normalized_langs,
                service=service,
                api_key=api_key,
                formats=normalized_formats,
                output_root=output_root,
                zip_output=True,
            )

            translations = summary.get('translations') or {}
            if not translations:
                raise RuntimeError('Translation failed: no outputs generated.')

            zip_path = summary.get('zip_path')
            download_path = None
            download_name = zip_download_name
            mimetype_value = 'application/zip'

            if zip_path and os.path.exists(zip_path):
                download_path = zip_path
            else:
                for translation in translations.values():
                    outputs = translation.get('outputs', {})
                    for path in outputs.values():
                        if path and os.path.exists(path):
                            download_path = path
                            mimetype_value = mimetypes.guess_type(path)[0] or 'application/octet-stream'
                            download_name = os.path.basename(path)
                            break
                    if download_path:
                        break

            if not download_path:
                raise RuntimeError('Translation completed but no output files were generated.')

            response = send_file(
                download_path,
                as_attachment=True,
                download_name=download_name,
                mimetype=mimetype_value
            )

            errors = summary.get('errors') or []
            if errors:
                response.headers['X-Translation-Warnings'] = '; '.join(errors)

            def cleanup():
                time.sleep(5)
                try:
                    if os.path.exists(input_path):
                        os.remove(input_path)
                except Exception:
                    pass

            cleanup_thread = threading.Thread(target=cleanup)
            cleanup_thread.daemon = True
            cleanup_thread.start()

            return response

        except ValueError as exc:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.isdir(output_root):
                shutil.rmtree(output_root, ignore_errors=True)
            return jsonify({'error': str(exc)}), 400

        except Exception as e:
            print(f"🚨 Translation Error: {str(e)}")  # Debug logging
            print(f"🚨 Error Type: {type(e).__name__}")  # Debug logging
            import traceback
            print(f"🚨 Full Traceback:")
            traceback.print_exc()  # Debug logging
            
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.isdir(output_root):
                shutil.rmtree(output_root, ignore_errors=True)
            return jsonify({'error': str(e)}), 500

    except Exception as e:
        print(f"🚨 Outer Exception: {str(e)}")  # Debug logging
        import traceback
        traceback.print_exc()  # Debug logging
        return jsonify({'error': str(e)}), 500

@app.route('/api/languages', methods=['GET'])
def get_languages():
    """Get all supported languages for Google Translate"""
    languages = {
        'af': 'Afrikaans', 'sq': 'Albanian', 'am': 'Amharic', 'ar': 'Arabic',
        'hy': 'Armenian', 'az': 'Azerbaijani', 'eu': 'Basque', 'be': 'Belarusian',
        'bn': 'Bengali', 'bs': 'Bosnian', 'bg': 'Bulgarian', 'ca': 'Catalan',
        'ceb': 'Cebuano', 'ny': 'Chichewa', 'zh-CN': 'Chinese (Simplified)',
        'zh-TW': 'Chinese (Traditional)', 'co': 'Corsican', 'hr': 'Croatian',
        'cs': 'Czech', 'da': 'Danish', 'nl': 'Dutch', 'en': 'English',
        'eo': 'Esperanto', 'et': 'Estonian', 'tl': 'Filipino', 'fi': 'Finnish',
        'fr': 'French', 'fy': 'Frisian', 'gl': 'Galician', 'ka': 'Georgian',
        'de': 'German', 'el': 'Greek', 'gu': 'Gujarati', 'ht': 'Haitian Creole',
        'ha': 'Hausa', 'haw': 'Hawaiian', 'iw': 'Hebrew', 'he': 'Hebrew',
        'hi': 'Hindi', 'hmn': 'Hmong', 'hu': 'Hungarian', 'is': 'Icelandic',
        'ig': 'Igbo', 'id': 'Indonesian', 'ga': 'Irish', 'it': 'Italian',
        'ja': 'Japanese', 'jw': 'Javanese', 'kn': 'Kannada', 'kk': 'Kazakh',
        'km': 'Khmer', 'ko': 'Korean', 'ku': 'Kurdish', 'ky': 'Kyrgyz',
        'lo': 'Lao', 'la': 'Latin', 'lv': 'Latvian', 'lt': 'Lithuanian',
        'lb': 'Luxembourgish', 'mk': 'Macedonian', 'mg': 'Malagasy', 'ms': 'Malay',
        'ml': 'Malayalam', 'mt': 'Maltese', 'mi': 'Maori', 'mr': 'Marathi',
        'mn': 'Mongolian', 'my': 'Myanmar (Burmese)', 'ne': 'Nepali', 'no': 'Norwegian',
        'or': 'Odia', 'ps': 'Pashto', 'fa': 'Persian', 'pl': 'Polish',
        'pt': 'Portuguese', 'pa': 'Punjabi', 'ro': 'Romanian', 'ru': 'Russian',
        'sm': 'Samoan', 'gd': 'Scots Gaelic', 'sr': 'Serbian', 'st': 'Sesotho',
        'sn': 'Shona', 'sd': 'Sindhi', 'si': 'Sinhala', 'sk': 'Slovak',
        'sl': 'Slovenian', 'so': 'Somali', 'es': 'Spanish', 'su': 'Sundanese',
        'sw': 'Swahili', 'sv': 'Swedish', 'tg': 'Tajik', 'ta': 'Tamil',
        'te': 'Telugu', 'th': 'Thai', 'tr': 'Turkish', 'uk': 'Ukrainian',
        'ur': 'Urdu', 'ug': 'Uyghur', 'uz': 'Uzbek', 'vi': 'Vietnamese',
        'cy': 'Welsh', 'xh': 'Xhosa', 'yi': 'Yiddish', 'yo': 'Yoruba',
        'zu': 'Zulu'
    }
    # Sort by language name
    sorted_langs = sorted(languages.items(), key=lambda x: x[1])
    return jsonify({'languages': dict(sorted_langs)})

@app.route('/api/status', methods=['GET'])
def status():
    return jsonify({
        'status': 'running',
        'services': ['google', 'deepl', 'openai'],
        'total_languages': 100
    })

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False, host="0.0.0.0", port=5001)

