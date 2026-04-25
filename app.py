import io
import os
import shutil
import subprocess
import zipfile
from datetime import datetime
from pathlib import Path
from threading import Lock
from tempfile import TemporaryDirectory

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from docx import Document
try:
    from docx2pdf import convert as docx_to_pdf
except Exception:
    docx_to_pdf = None
from docxcompose.composer import Composer
from pdf2docx import Converter
from pypdf import PdfMerger
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename


def get_int_env(name: str, default: int) -> int:
    value = os.environ.get(name)

    if value is None:
        return default

    try:
        parsed_value = int(value)
    except ValueError:
        return default

    return parsed_value if parsed_value > 0 else default


def get_bool_env(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default

    return value.strip().lower() in {"1", "true", "yes", "on"}


MAX_FILES_PER_REQUEST = get_int_env("MAX_FILES_PER_REQUEST", 30)
MAX_UPLOAD_SIZE_MB = get_int_env("MAX_UPLOAD_SIZE_MB", 250)
ALLOWED_PDF_EXTENSIONS = {"pdf"}
ALLOWED_WORD_EXTENSIONS = {"doc", "docx"}
ALLOWED_WORD_MERGE_EXTENSIONS = {"docx"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(32)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE_MB * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = get_bool_env("SESSION_COOKIE_SECURE", False)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

_status_lock = Lock()
_conversion_status = {
    "converted_files": 0,
    "conversion_jobs": 0,
    "merge_jobs": 0,
    "merged_source_files": 0,
}


def is_allowed_extension(filename: str, allowed_extensions: set[str]) -> bool:
    if "." not in filename:
        return False
    return filename.rsplit(".", 1)[1].lower() in allowed_extensions


def generate_unique_output_name(
    base_name: str,
    extension: str,
    used_names: set[str],
) -> str:
    output_name = f"{base_name}.{extension}"
    suffix = 1

    while output_name.lower() in used_names:
        output_name = f"{base_name}_{suffix}.{extension}"
        suffix += 1

    used_names.add(output_name.lower())
    return output_name


def make_zip_response(
    converted_files: list[Path],
    failed_files: list[tuple[str, str]],
    archive_prefix: str,
):
    archive_name = f"{archive_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for converted_file in converted_files:
            archive.write(converted_file, arcname=converted_file.name)

        if failed_files:
            report_lines = [
                "Some files could not be converted.",
                "",
                "Failure details:",
            ]
            report_lines.extend(f"- {name}: {reason}" for name, reason in failed_files)
            archive.writestr("conversion_report.txt", "\n".join(report_lines))

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=archive_name,
        mimetype="application/zip",
    )


def make_single_file_response(
    output_file: Path,
    download_name: str,
    mime_type: str,
):
    file_buffer = io.BytesIO(output_file.read_bytes())
    file_buffer.seek(0)

    return send_file(
        file_buffer,
        as_attachment=True,
        download_name=download_name,
        mimetype=mime_type,
    )


def update_conversion_status(
    converted_files: int = 0,
    conversion_jobs: int = 0,
    merge_jobs: int = 0,
    merged_source_files: int = 0,
) -> None:
    with _status_lock:
        _conversion_status["converted_files"] += max(converted_files, 0)
        _conversion_status["conversion_jobs"] += max(conversion_jobs, 0)
        _conversion_status["merge_jobs"] += max(merge_jobs, 0)
        _conversion_status["merged_source_files"] += max(merged_source_files, 0)


def get_conversion_status_snapshot() -> dict[str, int]:
    with _status_lock:
        snapshot = dict(_conversion_status)

    snapshot["total_jobs"] = snapshot["conversion_jobs"] + snapshot["merge_jobs"]
    return snapshot


def convert_pdf_to_docx(pdf_path: Path, docx_path: Path) -> None:
    converter = Converter(str(pdf_path))
    try:
        converter.convert(str(docx_path))
    finally:
        converter.close()


def merge_pdf_documents(pdf_paths: list[Path], merged_output_path: Path) -> None:
    merger = PdfMerger()
    try:
        for pdf_path in pdf_paths:
            merger.append(str(pdf_path))

        with merged_output_path.open("wb") as merged_file:
            merger.write(merged_file)
    finally:
        merger.close()


def merge_word_documents(docx_paths: list[Path], merged_output_path: Path) -> None:
    primary_document = Document(str(docx_paths[0]))
    composer = Composer(primary_document)

    for docx_path in docx_paths[1:]:
        composer.append(Document(str(docx_path)))

    composer.save(str(merged_output_path))


def convert_word_to_pdf_document(word_path: Path, pdf_path: Path) -> None:
    conversion_error = None

    if docx_to_pdf is not None:
        try:
            docx_to_pdf(str(word_path), str(pdf_path))
            if pdf_path.exists():
                return
            conversion_error = RuntimeError("Word conversion did not create a PDF output file.")
        except Exception as docx_to_pdf_error:  # noqa: BLE001
            conversion_error = docx_to_pdf_error

    soffice_path = shutil.which("soffice")
    if not soffice_path:
        raise RuntimeError(
            "Word to PDF conversion failed. Install Microsoft Word on Windows "
            "or LibreOffice (soffice) on hosted Linux environments."
        ) from conversion_error

    result = subprocess.run(
        [
            soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(word_path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )

    generated_pdf_path = pdf_path.parent / f"{word_path.stem}.pdf"
    conversion_logs = result.stderr.strip() or result.stdout.strip()

    if result.returncode != 0 or not generated_pdf_path.exists():
        raise RuntimeError(
            "LibreOffice conversion failed. "
            f"{conversion_logs or 'No additional details returned by soffice.'}"
        ) from conversion_error

    if generated_pdf_path != pdf_path:
        generated_pdf_path.replace(pdf_path)


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        max_files=MAX_FILES_PER_REQUEST,
        max_size_mb=MAX_UPLOAD_SIZE_MB,
        status=get_conversion_status_snapshot(),
    )


@app.route("/status", methods=["GET"])
def status():
    return jsonify(get_conversion_status_snapshot())


@app.route("/healthz", methods=["GET"])
def healthz():
    return jsonify({"status": "ok"})


@app.after_request
def set_response_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


@app.route("/convert", methods=["POST"])
@app.route("/convert-pdf-to-word", methods=["POST"])
def convert_pdf_to_word():
    uploaded_files = [
        uploaded_file
        for uploaded_file in request.files.getlist("pdf_files")
        if uploaded_file and uploaded_file.filename
    ]

    if not uploaded_files:
        flash("Please select at least one PDF file.")
        return redirect(url_for("index"))

    if len(uploaded_files) > MAX_FILES_PER_REQUEST:
        flash(f"You can upload up to {MAX_FILES_PER_REQUEST} files per conversion.")
        return redirect(url_for("index"))

    with TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        converted_files = []
        failed_files = []
        used_docx_names = set()

        for index, uploaded_file in enumerate(uploaded_files, start=1):
            original_filename = secure_filename(uploaded_file.filename or "")

            if not original_filename:
                failed_files.append((f"file_{index}", "Invalid filename"))
                continue

            if not is_allowed_extension(original_filename, ALLOWED_PDF_EXTENSIONS):
                failed_files.append((original_filename, "Only PDF files are allowed"))
                continue

            pdf_path = input_dir / f"{index:03d}_{original_filename}"
            uploaded_file.save(pdf_path)

            base_name = Path(original_filename).stem or f"converted_{index:03d}"
            docx_name = generate_unique_output_name(base_name, "docx", used_docx_names)

            docx_path = output_dir / docx_name
            try:
                convert_pdf_to_docx(pdf_path, docx_path)
                converted_files.append(docx_path)
            except Exception as conversion_error:  # noqa: BLE001
                failed_files.append((original_filename, str(conversion_error)))

        if not converted_files:
            if failed_files:
                preview = "; ".join(
                    f"{name}: {reason}" for name, reason in failed_files[:3]
                )
                flash(f"No files were converted. {preview}")
            else:
                flash("No valid PDF files were found.")
            return redirect(url_for("index"))

        update_conversion_status(
            converted_files=len(converted_files),
            conversion_jobs=1,
        )

        return make_zip_response(
            converted_files=converted_files,
            failed_files=failed_files,
            archive_prefix="converted_word_files",
        )


@app.route("/convert-word-to-pdf", methods=["POST"])
def convert_word_to_pdf():
    uploaded_files = [
        uploaded_file
        for uploaded_file in request.files.getlist("word_files")
        if uploaded_file and uploaded_file.filename
    ]

    if not uploaded_files:
        flash("Please select at least one Word file.")
        return redirect(url_for("index"))

    if len(uploaded_files) > MAX_FILES_PER_REQUEST:
        flash(f"You can upload up to {MAX_FILES_PER_REQUEST} files per conversion.")
        return redirect(url_for("index"))

    with TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        converted_files = []
        failed_files = []
        used_pdf_names = set()

        for index, uploaded_file in enumerate(uploaded_files, start=1):
            original_filename = secure_filename(uploaded_file.filename or "")

            if not original_filename:
                failed_files.append((f"file_{index}", "Invalid filename"))
                continue

            if not is_allowed_extension(original_filename, ALLOWED_WORD_EXTENSIONS):
                failed_files.append(
                    (original_filename, "Only DOC and DOCX files are allowed")
                )
                continue

            word_path = input_dir / f"{index:03d}_{original_filename}"
            uploaded_file.save(word_path)

            base_name = Path(original_filename).stem or f"converted_{index:03d}"
            pdf_name = generate_unique_output_name(base_name, "pdf", used_pdf_names)
            pdf_path = output_dir / pdf_name

            try:
                convert_word_to_pdf_document(word_path, pdf_path)
                converted_files.append(pdf_path)
            except Exception as conversion_error:  # noqa: BLE001
                failed_files.append((original_filename, str(conversion_error)))

        if not converted_files:
            if failed_files:
                preview = "; ".join(
                    f"{name}: {reason}" for name, reason in failed_files[:3]
                )
                flash(
                    "No files were converted. "
                    f"{preview}. On Windows, Word to PDF conversion requires Microsoft Word installed."
                )
            else:
                flash("No valid Word files were found.")
            return redirect(url_for("index"))

        update_conversion_status(
            converted_files=len(converted_files),
            conversion_jobs=1,
        )

        return make_zip_response(
            converted_files=converted_files,
            failed_files=failed_files,
            archive_prefix="converted_pdf_files",
        )


@app.route("/merge-pdf", methods=["POST"])
def merge_pdf_files():
    uploaded_files = [
        uploaded_file
        for uploaded_file in request.files.getlist("merge_pdf_files")
        if uploaded_file and uploaded_file.filename
    ]

    if not uploaded_files:
        flash("Please select PDF files to merge.")
        return redirect(url_for("index"))

    if len(uploaded_files) < 2:
        flash("Select at least two PDF files to merge.")
        return redirect(url_for("index"))

    if len(uploaded_files) > MAX_FILES_PER_REQUEST:
        flash(f"You can upload up to {MAX_FILES_PER_REQUEST} files per merge.")
        return redirect(url_for("index"))

    with TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        merge_sources: list[Path] = []

        for index, uploaded_file in enumerate(uploaded_files, start=1):
            original_filename = secure_filename(uploaded_file.filename or "")

            if not original_filename:
                flash("One of the selected files has an invalid filename.")
                return redirect(url_for("index"))

            if not is_allowed_extension(original_filename, ALLOWED_PDF_EXTENSIONS):
                flash(f"Only PDF files are allowed for merge. Invalid: {original_filename}")
                return redirect(url_for("index"))

            pdf_path = input_dir / f"{index:03d}_{original_filename}"
            uploaded_file.save(pdf_path)
            merge_sources.append(pdf_path)

        merged_filename = f"merged_pdf_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        merged_output_path = output_dir / merged_filename

        try:
            merge_pdf_documents(merge_sources, merged_output_path)
        except Exception as merge_error:  # noqa: BLE001
            flash(f"Unable to merge PDFs. {merge_error}")
            return redirect(url_for("index"))

        update_conversion_status(
            merge_jobs=1,
            merged_source_files=len(merge_sources),
        )

        return make_single_file_response(
            output_file=merged_output_path,
            download_name=merged_filename,
            mime_type="application/pdf",
        )


@app.route("/merge-word", methods=["POST"])
def merge_word_files():
    uploaded_files = [
        uploaded_file
        for uploaded_file in request.files.getlist("merge_word_files")
        if uploaded_file and uploaded_file.filename
    ]

    if not uploaded_files:
        flash("Please select Word files to merge.")
        return redirect(url_for("index"))

    if len(uploaded_files) < 2:
        flash("Select at least two Word files to merge.")
        return redirect(url_for("index"))

    if len(uploaded_files) > MAX_FILES_PER_REQUEST:
        flash(f"You can upload up to {MAX_FILES_PER_REQUEST} files per merge.")
        return redirect(url_for("index"))

    with TemporaryDirectory() as temp_dir:
        input_dir = Path(temp_dir) / "input"
        output_dir = Path(temp_dir) / "output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        merge_sources: list[Path] = []

        for index, uploaded_file in enumerate(uploaded_files, start=1):
            original_filename = secure_filename(uploaded_file.filename or "")

            if not original_filename:
                flash("One of the selected files has an invalid filename.")
                return redirect(url_for("index"))

            if not is_allowed_extension(original_filename, ALLOWED_WORD_MERGE_EXTENSIONS):
                flash(
                    "Word merge currently supports .docx only. "
                    f"Invalid file: {original_filename}"
                )
                return redirect(url_for("index"))

            docx_path = input_dir / f"{index:03d}_{original_filename}"
            uploaded_file.save(docx_path)
            merge_sources.append(docx_path)

        merged_filename = f"merged_word_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        merged_output_path = output_dir / merged_filename

        try:
            merge_word_documents(merge_sources, merged_output_path)
        except Exception as merge_error:  # noqa: BLE001
            flash(f"Unable to merge Word files. {merge_error}")
            return redirect(url_for("index"))

        update_conversion_status(
            merge_jobs=1,
            merged_source_files=len(merge_sources),
        )

        return make_single_file_response(
            output_file=merged_output_path,
            download_name=merged_filename,
            mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


@app.errorhandler(413)
def request_entity_too_large(_error):
    flash(
        "Upload is too large. Reduce file count or file size and try again. "
        f"Limit is {MAX_UPLOAD_SIZE_MB}MB per request."
    )
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(
        host=os.environ.get("FLASK_RUN_HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "5000")),
        debug=get_bool_env("FLASK_DEBUG", False),
    )