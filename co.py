import pandas as pd
import matplotlib.pyplot as plt
import os
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import datetime
import json
from reportlab.lib.fonts import addMapping

# Register Arabic fonts with full embedding
try:
    # Register the Arabic fonts
    pdfmetrics.registerFont(TTFont('Arabic', 'fonts/NotoSansArabic-Regular.ttf', validate=True))
    pdfmetrics.registerFont(TTFont('Arabic-Bold', 'fonts/NotoSansArabic-Bold.ttf', validate=True))
    
    # Add font mapping
    addMapping('Arabic', 0, 0, 'Arabic')  # normal
    addMapping('Arabic', 1, 0, 'Arabic-Bold')  # bold
    
    print("Arabic fonts registered successfully")
except Exception as e:
    print(f"Warning: Error loading Arabic fonts: {str(e)}")
    print("Using fallback fonts.")

# Load translations at module level
def load_translations():
    try:
        with open('translations.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading translations: {str(e)}")
        return {}

translations = load_translations()

def load_data(file_path):
    """Load data from Excel or CSV files"""
    try:
        print(f"\nDEBUG: Loading file: {file_path}")
        
        # Determine file type from extension
        file_extension = file_path.lower().split('.')[-1]
        
        if file_extension in ['xlsx', 'xls']:
            data = pd.read_excel(file_path, engine='openpyxl')
        elif file_extension == 'csv':
            # Try different encodings and delimiters
            encodings = ['utf-8', 'latin1', 'iso-8859-1']
            delimiters = [',', ';', '\t']
            
            for encoding in encodings:
                for delimiter in delimiters:
                    try:
                        data = pd.read_csv(file_path, encoding=encoding, sep=delimiter)
                        print(f"DEBUG: Successfully read CSV with encoding {encoding} and delimiter {delimiter}")
                        break
                    except Exception as e:
                        continue
                if 'data' in locals():
                    break
                    
            if 'data' not in locals():
                raise Exception("Could not read CSV file with any combination of encoding and delimiter")
        else:
            raise Exception(f"Unsupported file format: {file_extension}")
            
        print("DEBUG: File loaded successfully")
        print(f"DEBUG: Found {len(data)} rows and {len(data.columns)} columns")
        return data
        
    except Exception as e:
        print(f"\nERROR in load_data:")
        print(f"Error message: {str(e)}")
        raise Exception(f"Error loading file {file_path}: {str(e)}")

def clean_data(data):
    """Clean and prepare the data"""
    try:
        # Debug: Print initial data info
        print("\nDEBUG: Data Info:")
        print("Shape:", data.shape)
        print("Columns:", data.columns.tolist())
        print("\nFirst few rows:")
        print(data.head())
        
        # Map expected column names (handles both French and English)
        column_mapping = {
            'Profil': ['Profil', 'Profile', 'PROFIL'],
            'Qté': ['Qté', 'Qty', 'Quantité', 'QTE', 'QTÉ'],
            'Long.': ['Long.', 'Length', 'LONG.'],
            'Poids': ['Poids', 'Weight', 'POIDS']
        }
        
        # Debug: Print column search process
        print("\nDEBUG: Searching for columns:")
        actual_columns = {}
        for expected, possibilities in column_mapping.items():
            print(f"\nLooking for {expected} in possibilities: {possibilities}")
            found = False
            for col in possibilities:
                if col in data.columns:
                    actual_columns[expected] = col
                    print(f"Found: {col}")
                    found = True
                    break
            if not found:
                print(f"WARNING: Could not find column {expected}")
                print(f"Available columns: {data.columns.tolist()}")
                raise Exception(f"Missing required column {expected}")
        
        # Extract and rename columns
        data_cleaned = data[[
            actual_columns['Profil'],
            actual_columns['Qté'],
            actual_columns['Long.'],
            actual_columns['Poids']
        ]].copy()
        
        # Rename columns
        data_cleaned.columns = ['Profil', 'Qté', 'Long.', 'Poids']
        
        # Remove rows with NaN values and 'Total' rows
        data_cleaned = data_cleaned.dropna(subset=['Profil', 'Qté', 'Long.'])
        data_cleaned = data_cleaned[~data_cleaned['Profil'].str.contains('Total', na=False)]
        
        # Convert to proper types
        data_cleaned['Qté'] = pd.to_numeric(data_cleaned['Qté'], errors='coerce').fillna(1).astype(int)
        data_cleaned['Long.'] = pd.to_numeric(data_cleaned['Long.'], errors='coerce').fillna(0).astype(int)
        data_cleaned['Poids'] = pd.to_numeric(data_cleaned['Poids'], errors='coerce').fillna(0).astype(float)
        
        # Calculate Pds Tot
        data_cleaned['Pds Tot'] = data_cleaned['Qté'] * data_cleaned['Poids']
        
        print("\nDEBUG: Final cleaned data with calculated total weight:")
        print(data_cleaned.head())
        
        return data_cleaned
        
    except Exception as e:
        print("\nERROR in clean_data:")
        print(f"Error message: {str(e)}")
        print("Full data columns:", data.columns.tolist())
        raise Exception(f"Error cleaning data: {str(e)}")

def get_stock_length(profile, settings_df, default_length):
    """Get stock length for a profile from settings DataFrame"""
    length_row = settings_df[settings_df['Profile'] == profile]
    if not length_row.empty:
        return length_row['Stock Length'].iloc[0]
    return default_length

def optimize_cutting(data, settings_df, default_length):
    results = []
    
    # Process each unique profile
    for profile in settings_df['Profile'].unique():
        # Get stock length for this profile
        stock_length = settings_df[settings_df['Profile'] == profile]['Stock Length'].iloc[0]
        
        # Get all pieces for this profile
        profile_data = data[data['Profil'] == profile]
        
        pieces = []
        if not profile_data.empty:
            # Collect all pieces with their quantities
            for _, row in profile_data.iterrows():
                pieces.extend([int(row['Long.'])] * int(row['Qté']))
        
        if pieces:  # Only process if we have pieces to cut
            pieces.sort(reverse=True)  # Sort pieces in descending order
            stock_used = []
            current_stock = stock_length
            current_pieces = []

            for piece in pieces:
                if piece <= current_stock:
                    current_pieces.append(piece)
                    current_stock -= piece
                else:
                    if current_pieces:
                        stock_used.append((current_pieces, stock_length - current_stock))
                    current_pieces = [piece]
                    current_stock = stock_length - piece

            if current_pieces:
                stock_used.append((current_pieces, stock_length - current_stock))
            
            results.append((profile, stock_length, stock_used))
    
    return results

def draw_cutting_plan(results, save_path):
    """Draw cutting plan and save to AppData"""
    try:
        fig, axs = plt.subplots(len(results), figsize=(10, len(results) * 2))
        
        if len(results) == 1:
            axs = [axs]
        
        for i, (profile, stock_length, stock_used) in enumerate(results):
            for j, (pieces, total_used) in enumerate(stock_used):
                axs[i].barh(y=j, width=stock_length, color='grey', edgecolor='black')
                start = 0
                for piece in pieces:
                    axs[i].barh(y=j, width=piece, left=start, color='blue', edgecolor='black')
                    start += piece
                axs[i].set_xlim(0, stock_length)
                axs[i].set_title(f'Profile {profile}: Piece {j+1}: Total Length Used = {total_used} mm, Remaining = {stock_length - total_used} mm')
                axs[i].axis('off')

        plt.tight_layout()
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        plt.savefig(save_path)
        plt.close()
        
    except Exception as e:
        print(f"Error drawing cutting plan: {str(e)}")
        raise

def export_to_excel(results, output_path, image_path, language="fr"):
    """Export cutting plan to Excel"""
    try:
        # Get translations with fallback
        if language not in translations:
            language = "en"
            
        # Get translations for cutting plan
        t = translations[language]
        if not t:
            raise ValueError(f"Translations not found for language: {language}")
        
        # Create column headers dictionary
        columns = {
            "profile": t.get("Profile", "Profile"),
            "stock_length": t.get("Stock Length", "Stock Length"),
            "cut_index": t.get("Cut Index", "Cut Index"),
            "pieces_cut": t.get("Pieces Cut", "Pieces Cut"),
            "total_length": t.get("Total Length Used", "Total Length Used"),
            "remaining": t.get("Remaining Length", "Remaining Length")
        }
        
        data_to_export = []
        for profile, stock_length, stock_used in results:
            for j, (pieces, remaining) in enumerate(stock_used):
                data_to_export.append({
                    columns["profile"]: profile,
                    columns["stock_length"]: stock_length,
                    columns["cut_index"]: j+1,
                    columns["pieces_cut"]: pieces,
                    columns["total_length"]: sum(pieces),
                    columns["remaining"]: remaining
                })
        
        df = pd.DataFrame(data_to_export)
        
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=t.get("Cutting Plan", "Cutting Plan"), index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[t.get("Cutting Plan", "Cutting Plan")]
        
        # Add RTL support for Arabic
        if language == "ar":
            worksheet.right_to_left()
            
        worksheet.insert_image('H2', image_path)
        writer._save()
        
    except Exception as e:
        print(f"Error exporting to Excel: {str(e)}")
        raise

def calculate_waste_percentage(results):
    """Calculate waste percentage for each profile"""
    waste_stats = {}
    for profile, stock_length, stock_used in results:
        total_stock = len(stock_used) * stock_length
        used_length = sum(sum(pieces) for pieces, _ in stock_used)
        waste_percentage = ((total_stock - used_length) / total_stock) * 100
        waste_stats[profile] = {
            'waste_percentage': round(waste_percentage, 2),
            'total_stock': total_stock,
            'used_length': used_length
        }
    return waste_stats

def calculate_weight_stats(data_df):
    """Calculate weight statistics for each profile"""
    weight_stats = {}
    
    for profile in data_df['Profil'].unique():
        profile_data = data_df[data_df['Profil'] == profile]
        total_weight = profile_data['Pds Tot'].sum()
        weight_stats[profile] = round(total_weight, 3)
    
    return weight_stats

def export_invoice_pdf(results, weight_stats, total_weight, adjusted_weight, steel_price, weight_error, output_path, language="fr"):
    """Export a PDF invoice"""
    try:
        # Get translations for the current language
        if language not in translations:
            language = "en"
            
        t = translations[language].get("invoice", {})
        if not t:
            raise ValueError(f"Invoice translations not found for language: {language}")
        
        # Create PDF document with font embedding
        pdf_path = output_path.replace('.xlsx', '_facture.pdf')
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=A4,
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30,
            encoding='UTF-8'  # Ensure UTF-8 encoding
        )
        
        # Styles with RTL support
        styles = getSampleStyleSheet()
        if language == "ar":
            # Arabic styles with explicit font encoding
            style_title = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                fontName='Arabic-Bold',
                alignment=2,  # Right alignment for Arabic
                wordWrap='RTL',
                spaceAfter=30,
                encoding='UTF-8'
            )
            normal_style = ParagraphStyle(
                'ArabicNormal',
                parent=styles['Normal'],
                fontName='Arabic',
                alignment=2,  # Right alignment for Arabic
                wordWrap='RTL',
                encoding='UTF-8'
            )
            
            # Create table style with Arabic font
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Arabic-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Arabic'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, -2), (-1, -1), colors.lightgrey),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, -2), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ])
        else:
            # Default styles for other languages
            style_title = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=30,
                alignment=1
            )
            normal_style = styles['Normal']
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, -2), (-1, -1), colors.lightgrey),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('ALIGN', (0, -2), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ])
        
        # Content elements
        elements = []
        
        # Add title and date with proper RTL handling
        elements.append(Paragraph(t["title"], style_title))
        date_text = f"{t['date']}: {datetime.datetime.now().strftime('%d/%m/%Y')}"
        elements.append(Paragraph(date_text, normal_style))
        elements.append(Spacer(1, 20))
        
        # Prepare table data using translations
        headers = [
            t["headers"]["type"],
            t["headers"]["weight_per_meter"],
            t["headers"]["stock_length"],
            t["headers"]["quantity"],
            t["headers"]["total_weight"],
            t["headers"]["percentage"],
            t["headers"]["total_price"]
        ]
        
        data = [headers]
        
        # Add data rows
        for result in results:
            profile = result[0]
            stock_length = result[1]
            stock_used = result[2]
            
            qty = sum(len(pieces) for pieces, _ in stock_used)
            profile_weight = weight_stats.get(profile, 0)
            weight_per_unit = profile_weight/qty if qty > 0 else 0
            percentage = profile_weight / total_weight if total_weight > 0 else 0
            price = profile_weight * steel_price
            
            row = [
                profile,
                f"{weight_per_unit:.3f}",
                f"{stock_length}",
                f"{qty}",
                f"{profile_weight:.3f}",
                f"{percentage:.1%}",
                f"{price:.3f}"
            ]
            data.append(row)
        
        # Add totals
        total_price = total_weight * steel_price
        adjusted_price = adjusted_weight * steel_price
        adjusted_percentage = 1 + (weight_error/100)
        
        # Total row using translations
        data.append([
            t["total"],
            "",
            "",
            "",
            f"{total_weight:.3f}",
            "100%",
            f"{total_price:.3f}"
        ])
        
        # Adjusted total row using translations
        data.append([
            t["total_adjusted"].format(weight_error),
            "",
            "",
            "",
            f"{adjusted_weight:.3f}",
            f"{adjusted_percentage:.0%}",
            f"{adjusted_price:.3f}"
        ])
        
        # Create table
        table = Table(data, colWidths=[35*mm, 25*mm, 25*mm, 20*mm, 30*mm, 25*mm, 30*mm])
        
        # Add style to table
        table.setStyle(table_style)
        
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        print(f"Invoice PDF exported to: {pdf_path}")
        
    except Exception as e:
        print(f"Error exporting PDF invoice: {str(e)}")
        raise

def export_invoice_excel(results, weight_stats, total_weight, adjusted_weight, steel_price, weight_error, output_path, language="fr"):
    """Export invoice to Excel"""
    try:
        # Get translations with fallback
        if language not in translations:
            language = "en"
            
        t = translations[language].get("invoice", {})
        if not t:
            raise ValueError(f"Invoice translations not found for language: {language}")
        
        # Create Excel workbook
        excel_path = output_path.replace('.xlsx', '_facture.xlsx')
        workbook = xlsxwriter.Workbook(excel_path)
        worksheet = workbook.add_worksheet("Facture")
        
        # Add RTL support for Arabic
        if language == "ar":
            worksheet.right_to_left()
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D3D3D3',
            'border': 1
        })
        
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        number_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.000'
        })
        
        percent_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '0.0%'
        })
        
        # Set column widths
        worksheet.set_column('A:A', 15)  # Type
        worksheet.set_column('B:B', 15)  # Weight per unit
        worksheet.set_column('C:C', 12)  # Stock length
        worksheet.set_column('D:D', 10)  # Quantity
        worksheet.set_column('E:E', 15)  # Total weight
        worksheet.set_column('F:F', 12)  # Percentage
        worksheet.set_column('G:G', 15)  # Total price
        
        # Write title and date
        title = translations[language]["invoice"]["title"]
        date = translations[language]["invoice"]["date"]
        worksheet.write(0, 0, title, header_format)
        worksheet.write(1, 0, f"{date}: {datetime.datetime.now().strftime('%d/%m/%Y')}", cell_format)
        
        # Write headers
        headers = translations[language]["invoice"]["headers"]
        row = 3
        for col, header in enumerate(headers.values()):
            worksheet.write(row, col, header, header_format)
        
        # Write data
        row = 4
        for result in results:
            profile = result[0]
            stock_length = result[1]
            stock_used = result[2]
            
            qty = sum(len(pieces) for pieces, _ in stock_used)
            profile_weight = weight_stats.get(profile, 0)
            weight_per_unit = profile_weight/qty if qty > 0 else 0
            percentage = profile_weight / total_weight if total_weight > 0 else 0
            price = profile_weight * steel_price
            
            worksheet.write(row, 0, profile, cell_format)
            worksheet.write(row, 1, f"{weight_per_unit:.3f}", cell_format)
            worksheet.write(row, 2, stock_length, cell_format)
            worksheet.write(row, 3, qty, cell_format)
            worksheet.write(row, 4, profile_weight, cell_format)
            worksheet.write(row, 5, f"{percentage:.1%}", cell_format)
            worksheet.write(row, 6, f"{price:.3f}", cell_format)
            row += 1
        
        # Write totals
        total_price = total_weight * steel_price
        adjusted_price = adjusted_weight * steel_price
        adjusted_percentage = 1 + (weight_error/100)  # Calculate actual percentage
        
        worksheet.write(row, 0, translations[language]["invoice"]["total"], header_format)
        worksheet.write(row, 1, "", header_format)
        worksheet.write(row, 2, "", header_format)
        worksheet.write(row, 3, "", header_format)
        worksheet.write(row, 4, f"{total_weight:.3f}", cell_format)
        worksheet.write(row, 5, "100%", cell_format)
        worksheet.write(row, 6, f"{total_price:.3f}", cell_format)
        row += 1
        
        worksheet.write(row, 0, f"{translations[language]['invoice']['total_adjusted']} ({weight_error}%)", header_format)
        worksheet.write(row, 1, "", header_format)
        worksheet.write(row, 2, "", header_format)
        worksheet.write(row, 3, "", header_format)
        worksheet.write(row, 4, f"{adjusted_weight:.3f}", cell_format)
        worksheet.write(row, 5, f"{adjusted_percentage:.0%}", cell_format)
        worksheet.write(row, 6, f"{adjusted_price:.3f}", cell_format)
        row += 1
        
        # Save workbook
        workbook.close()
        print(f"Invoice Excel exported to: {excel_path}")
        
    except Exception as e:
        print(f"Error exporting invoice to Excel: {str(e)}")
        raise

def main(data_df, settings_df, input_filename, default_length, weight_error, steel_price, language="fr"):
    try:
        # Get base filename without extension
        base_filename = os.path.splitext(os.path.basename(input_filename))[0]
        
        # Get AppData path and create output directory with file name
        app_data = os.path.join(os.getenv('APPDATA'), 'Cutting Optimizer Pro')
        output_dir = os.path.join(app_data, 'output', base_filename)
        os.makedirs(output_dir, exist_ok=True)
        
        # Create output file paths
        output_excel_path = os.path.join(output_dir, f'{base_filename}_cutting_plan.xlsx')
        output_invoice_excel = os.path.join(output_dir, f'{base_filename}_facture.xlsx')
        output_image_path = os.path.join(output_dir, f'{base_filename}_cutting_plan.png')
        output_pdf_path = os.path.join(output_dir, f'{base_filename}_facture.pdf')
        
        # Clean data
        data_cleaned = clean_data(data_df)
        
        # Run optimization
        results = optimize_cutting(data_cleaned, settings_df, default_length)
        
        # Calculate statistics
        waste_stats = calculate_waste_percentage(results)
        weight_stats = calculate_weight_stats(data_cleaned)
        
        # Calculate total and adjusted weights
        total_weight = sum(weight_stats.values())
        adjusted_weight = total_weight * (1 + weight_error/100)
        total_price = adjusted_weight * steel_price
        
        # Draw cutting plan
        draw_cutting_plan(results, output_image_path)
        
        # Export to Excel files
        export_to_excel(results, output_excel_path, output_image_path, language)
        export_invoice_excel(
            results,
            weight_stats,
            total_weight,
            adjusted_weight,
            steel_price,
            weight_error,
            output_invoice_excel,
            language
        )
        
        # Export invoice PDF
        export_invoice_pdf(
            results,
            weight_stats,
            total_weight,
            adjusted_weight,
            steel_price,
            weight_error,
            output_pdf_path,
            language
        )
        
        return {
            'waste': waste_stats,
            'weight': {
                'total': round(total_weight, 3),
                'adjusted': round(adjusted_weight, 3),
                'price': round(total_price, 2)
            }
        }
        
    except Exception as e:
        print(f"Error in main function: {str(e)}")
        raise
