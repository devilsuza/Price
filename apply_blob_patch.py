from pathlib import Path

app_file = Path("app.py")

if not app_file.exists():
    raise SystemExit("ERROR: app.py repository में नहीं मिली")

code = app_file.read_text(encoding="utf-8")

old_storage = '''app = Flask(__name__)
UPLOAD_FOLDER = "temp_uploads"
DB_NAME = "store_prices.db"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)'''

new_storage = '''app = Flask(__name__)

# ============================================================
# VERCEL BLOB PERMANENT SQLITE STORAGE
# ============================================================

import tempfile
import threading

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IS_VERCEL = bool(os.environ.get("VERCEL"))

RUNTIME_DIR = (
    tempfile.gettempdir()
    if IS_VERCEL
    else BASE_DIR
)

UPLOAD_FOLDER = os.path.join(
    RUNTIME_DIR,
    "temp_uploads"
)

DB_NAME = os.path.join(
    RUNTIME_DIR,
    "store_prices.db"
)

BLOB_DB_PATH = os.environ.get(
    "BLOB_DB_PATH",
    "fresmart/store_prices.db"
)

BLOB_ACCESS = os.environ.get(
    "BLOB_ACCESS",
    "private"
).strip().lower()

_BLOB_LOCK = threading.RLock()

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def blob_is_configured():
    return bool(
        os.environ.get("BLOB_STORE_ID")
        or os.environ.get("BLOB_READ_WRITE_TOKEN")
        or os.environ.get("VERCEL_BLOB_READ_WRITE_TOKEN")
    )


def restore_database_from_blob():
    """Restore database from Vercel Blob when a new instance starts."""

    if not blob_is_configured():
        print(
            "[BLOB] Store not connected. "
            "Using normal local SQLite mode."
        )
        return False

    restore_file = DB_NAME + ".restore"

    try:
        from vercel.blob import sync as blob

        result = blob.get_blob(
            BLOB_DB_PATH,
            access=BLOB_ACCESS
        )

        with open(restore_file, "wb") as file:
            file.write(result.body)

        with sqlite3.connect(restore_file) as connection:
            check = connection.execute(
                "PRAGMA quick_check"
            ).fetchone()

            if not check or check[0] != "ok":
                raise RuntimeError(
                    "Downloaded SQLite database is not valid"
                )

        os.replace(restore_file, DB_NAME)

        print(
            f"[BLOB] Database restored from: "
            f"{BLOB_DB_PATH}"
        )

        return True

    except FileNotFoundError:
        print(
            "[BLOB] Cloud database does not exist yet. "
            "It will be created after the first update."
        )
        return False

    except Exception as error:
        print(
            f"[BLOB] Database restore warning: {error}"
        )

        try:
            if os.path.exists(restore_file):
                os.remove(restore_file)
        except OSError:
            pass

        return False


def backup_database_to_blob(connection=None):
    """Create a consistent SQLite snapshot and save it to Blob."""

    if not blob_is_configured():
        print(
            "[BLOB] Backup skipped: "
            "Blob Store is not connected."
        )
        return False

    if not os.path.exists(DB_NAME):
        print(
            "[BLOB] Backup skipped: "
            "Database does not exist."
        )
        return False

    snapshot_file = DB_NAME + ".snapshot"

    with _BLOB_LOCK:
        try:
            owns_connection = connection is None

            source_connection = (
                connection
                if connection is not None
                else sqlite3.connect(
                    DB_NAME,
                    timeout=30
                )
            )

            try:
                with sqlite3.connect(
                    snapshot_file,
                    timeout=30
                ) as destination_connection:

                    source_connection.backup(
                        destination_connection
                    )

            finally:
                if owns_connection:
                    source_connection.close()

            from vercel.blob import sync as blob

            snapshot_size = os.path.getsize(
                snapshot_file
            )

            with open(snapshot_file, "rb") as file:
                blob.put_blob(
                    BLOB_DB_PATH,
                    file,
                    access=BLOB_ACCESS,
                    content_length=snapshot_size,
                    content_type="application
