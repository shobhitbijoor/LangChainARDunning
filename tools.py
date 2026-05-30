@tool
def send_invoice_to_customer(invoice_number: str) -> str:
    """
    Sends a copy of an invoice PDF to the customer by copying it from the input invoices
    directory to the output directory. Use this tool when a customer has requested a copy
    of their invoice. The invoice_number should match the PDF filename (without extension).
    """
    source_path = INPUT_INVOICES_DIR / f"{invoice_number}.pdf"
    destination_path = OUTPUT_DIR / f"{invoice_number}.pdf"

    if not source_path.exists():
        return (
            f"Error: Invoice PDF for '{invoice_number}' not found at '{source_path}'. "
            f"Please verify the invoice number is correct."
        )

    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(source_path), str(destination_path))
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return (
            f"[{timestamp}] Successfully sent invoice '{invoice_number}'. "
            f"Copy saved to '{destination_path}'."
        )
    except Exception as e:
        return f"Error sending invoice '{invoice_number}': {str(e)}"


@tool
def update_inquiry_status_database(
    invoice_number: str,
    customer_name: str,
    customer_email: str,
    inquiry_type: str,
    promise_date: str = "",
    notes: str = "",
) -> str:
    """
    Records a processed customer inquiry into the SQLite database. Use this tool after
    handling any customer email to log the outcome. The inquiry_type must be either
    'invoice_request' (customer asked for a copy of their invoice) or 'promise_date'
    (customer provided a date by which they promise to pay). Optionally supply a
    promise_date (YYYY-MM-DD) and any notes about the interaction.
    """
    if inquiry_type not in ("invoice_request", "promise_date"):
        return (
            f"Error: inquiry_type must be 'invoice_request' or 'promise_date', "
            f"got '{inquiry_type}'."
        )

    try:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS invoice_inquiry (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                customer_name  TEXT NOT NULL,
                customer_email TEXT NOT NULL,
                inquiry_type   TEXT NOT NULL,
                promise_date   TEXT,
                status         TEXT DEFAULT 'processed',
                email_file     TEXT,
                notes          TEXT,
                processed_at   TEXT NOT NULL
            )
            """
        )

        processed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute(
            """
            INSERT INTO invoice_inquiry
                (invoice_number, customer_name, customer_email, inquiry_type,
                 promise_date, status, email_file, notes, processed_at)
            VALUES (?, ?, ?, ?, ?, 'processed', NULL, ?, ?)
            """,
            (
                invoice_number,
                customer_name,
                customer_email,
                inquiry_type,
                promise_date if promise_date else None,
                notes if notes else None,
                processed_at,
            ),
        )

        conn.commit()
        conn.close()

        return (
            f"[{processed_at}] Successfully recorded inquiry in database. "
            f"Invoice: '{invoice_number}', Customer: '{customer_name}', "
            f"Type: '{inquiry_type}', Status: 'processed'."
        )
    except Exception as e:
        return f"Error updating inquiry database for invoice '{invoice_number}': {str(e)}"


@tool
def update_erp(
    invoice_number: str,
    promise_date: str,
    customer_name: str = "",
    notes: str = "",
) -> str:
    """
    Updates the ERP Excel file with a customer's payment promise date. Use this tool
    when a customer has provided a date by which they promise to make payment. The
    tool finds the matching invoice row by invoice_number, sets the Promise Date column
    to promise_date (YYYY-MM-DD), and changes the Status to 'Promise Received'.
    Optionally updates the Notes column if notes are provided.
    """
    if not ERP_EXCEL_PATH.exists():
        return (
            f"Error: ERP Excel file not found at '{ERP_EXCEL_PATH}'. "
            f"Please verify the file path."
        )

    try:
        df = pd.read_excel(str(ERP_EXCEL_PATH))

        mask = df["Invoice Number"].astype(str).str.strip() == str(invoice_number).strip()

        if not mask.any():
            return (
                f"Error: Invoice '{invoice_number}' not found in ERP file '{ERP_EXCEL_PATH}'. "
                f"Please verify the invoice number."
            )

        df.loc[mask, "Promise Date"] = promise_date
        df.loc[mask, "Status"] = "Promise Received"

        if notes:
            df.loc[mask, "Notes"] = notes

        df.to_excel(str(ERP_EXCEL_PATH), index=False)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return (
            f"[{timestamp}] Successfully updated ERP for invoice '{invoice_number}'. "
            f"Promise Date set to '{promise_date}', Status set to 'Promise Received'."
        )
    except Exception as e:
        return f"Error updating ERP for invoice '{invoice_number}': {str(e)}"
