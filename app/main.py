from flask import Flask, render_template, request, Request
import os
from werkzeug.utils import secure_filename
from .modules import process_and_merge_workbook, get_suffix, load_acronyms


app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
STORAGE_FOLDER = os.path.join(os.path.dirname(__file__), "storage")
ALLOWED_EXTENSIONS = {'.xlsx', '.xlsm'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['STORAGE_FOLDER'] = STORAGE_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def _allowed_file(file_name: str) -> bool:
    return get_suffix(file_name).lower() in ALLOWED_EXTENSIONS


def _handle_upload(request: Request) -> tuple[str, str]:
    if 'file' not in request.files:
        raise ValueError("Arquivo não encontrado")
    
    file = request.files['file']

    if file.filename == '' or not file:
        raise ValueError("Arquivo inválido")
    
    if not _allowed_file(file.filename):
        raise ValueError(f"Arquivo com extensão inválida. Utilize {ALLOWED_EXTENSIONS}.")
    
    file_name = secure_filename(file.filename)
    uploaded_file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    file.save(uploaded_file_path)

    return uploaded_file_path, file_name


@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            uploaded_file_path, file_name = _handle_upload(request)

            acronym = request.form['acronym']
            tariff_process = request.form['tariff_process']
            process_date_str = request.form['process_date_str']

            try:
                process_and_merge_workbook(
                    uploaded_file_path=uploaded_file_path,
                    acronym=acronym,
                    tariff_process=tariff_process,
                    process_date_str=process_date_str
                )

                if os.path.exists(uploaded_file_path):
                    os.remove(uploaded_file_path)

                return f"Arquivo {file_name} processado com sucesso!"
            except Exception as error:
                return f"Ocorreu um erro ao processar o arquivo: {str(error)}"
        except Exception as error:
            return f"{str(error)}"

    acronyms_concessionaria = load_acronyms("Concessionária")
    acronyms_permissionaria = load_acronyms("Permissionária")
    
    return render_template(
        'upload.html',
        acronyms_concessionaria=acronyms_concessionaria,
        acronyms_permissionaria=acronyms_permissionaria
    )


# if __name__ == "__main__":
#     app.run(debug=True)