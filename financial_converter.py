import os
import sys
import time
import pandas as pd
import xlsxwriter
from docling.document_converter import DocumentConverter
from rich.console import Console
from rich.progress import track

# Initialize Rich Console for pretty printing
console = Console()


def create_formatted_excel(conv_res, output_path):
    """
    Creates a professionally formatted Excel file from Docling conversion results.
    """
    doc = conv_res.document
    tables = doc.tables

    if not tables:
        console.print("[yellow]No tables found in the document.[/yellow]")
        return

    console.print(f"[cyan]Writing {len(tables)} tables to Excel...[/cyan]")

    # Create a Pandas Excel writer using XlsxWriter as the engine
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Define Formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',  # Light green background
                'border': 1,
                'font_size': 11
            })

            cell_format = workbook.add_format({
                'text_wrap': False,
                'valign': 'top',
                'border': 1,
                'font_size': 10
            })

            number_format = workbook.add_format({
                'border': 1,
                'font_size': 10,
                'num_format': '#,##0.00'  # Standard financial number format
            })

            for i, table in track(enumerate(tables), description="Processing tables...", total=len(tables)):
                # Export table to DataFrame
                # Docling handles the complex structure extraction here
                df = table.export_to_dataframe(doc=doc)

                # Clean column names (flatten multi-index if necessary, though Docling usually handles this)
                df.columns = [str(c).strip() for c in df.columns]

                # Generate sheet name (Excel has 31 char limit)
                sheet_name = f"Table_{i + 1}"

                # Write the DataFrame to a worksheet
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0, header=False)

                worksheet = writer.sheets[sheet_name]

                # --- Beautification Logic ---

                # 1. Write the Header Row with Formatting
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, value, header_format)

                # 2. Auto-fit columns and apply cell formatting
                # (Iterate through rows to find max width and detect data types)
                for col_num, col_name in enumerate(df.columns):
                    max_len = len(str(col_name))

                    # Iterate through rows for width calculation and formatting
                    for row_num, value in enumerate(df[col_name]):
                        # Attempt to detect numbers for specific formatting
                        is_number = False
                        try:
                            # Check if it's a number but not a string that looks like a number
                            if isinstance(value, (int, float)):
                                is_number = True
                                # Write with number format
                                worksheet.write(row_num + 1, col_num, value, number_format)
                            else:
                                # Write with text format
                                worksheet.write(row_num + 1, col_num, str(value), cell_format)
                        except:
                            pass

                        # Update max length for auto-width
                        try:
                            cell_len = len(str(value))
                            if cell_len > max_len:
                                max_len = cell_len
                        except:
                            pass

                    # Set column width (add a little padding)
                    # Cap width at 50 characters for readability
                    width = min(max_len + 2, 50)
                    worksheet.set_column(col_num, col_num, width)

                # 3. Freeze the Header Row
                worksheet.freeze_panes(1, 0)

        console.print(f"[bold green]Success![/bold green] File saved to: {output_path}")

    except Exception as e:
        console.print(f"[bold red]Error writing Excel file:[/bold red] {e}")


def process_pdf(pdf_path):
    """
    Main processing function.
    """
    if not os.path.exists(pdf_path):
        console.print(f"[red]File not found: {pdf_path}[/red]")
        return

    console.print(f"[bold blue]Analyzing Document:[/bold blue] {os.path.basename(pdf_path)}")

    # Initialize Docling Converter
    # Note: You can configure options here if needed, e.g., enabling OCR
    doc_converter = DocumentConverter()

    start_time = time.time()

    try:
        # Convert the document
        # This step does the heavy AI lifting
        conv_res = doc_converter.convert(pdf_path)

        processing_time = time.time() - start_time
        console.print(f"Docling extraction took {processing_time:.2f} seconds.")

        # Prepare output filename
        base_name = os.path.splitext(pdf_path)[0]
        output_path = f"{base_name}_structured.xlsx"

        # Create the beautified Excel file
        create_formatted_excel(conv_res, output_path)

    except Exception as e:
        console.print(f"[bold red]Processing Failed:[/bold red] {e}")


if __name__ == "__main__":
    # Check if file path is provided via command line
    if len(sys.argv) < 2:
        console.print("[yellow]Usage:[/yellow] python financial_converter.py <path_to_pdf>")
        console.print("Or drag and drop a PDF file onto this script.")
        input("Press Enter to exit...")
        sys.exit(1)

    # Get the file path
    input_file = sys.argv[1]
    process_pdf(input_file)