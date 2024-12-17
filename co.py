import pandas as pd
import matplotlib.pyplot as plt
import os
import xlsxwriter

def load_data(file_path):
    """Load data from Excel files"""
    try:
        print(f"\nDEBUG: Loading file: {file_path}")
        data = pd.read_excel(file_path, engine='openpyxl')
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
    plt.savefig(save_path)
    plt.close()

def export_to_excel(results, output_path, image_path, language="en"):
    # Translation dictionary
    translations = {
        "en": {
            "Profile": "Profile",
            "Stock Length": "Stock Length",
            "Cut Index": "Cut Index",
            "Pieces Cut": "Pieces Cut",
            "Total Length Used": "Total Length Used",
            "Remaining Length": "Remaining Length",
            "Cutting Plan": "Cutting Plan"
        },
        "fr": {
            "Profile": "Profil",
            "Stock Length": "Longueur Stock",
            "Cut Index": "Index Coupe",
            "Pieces Cut": "Pièces Coupées",
            "Total Length Used": "Longueur Totale Utilisée",
            "Remaining Length": "Longueur Restante",
            "Cutting Plan": "Plan de Découpe"
        }
    }
    
    t = translations[language]
    
    data_to_export = []
    for profile, stock_length, stock_used in results:
        for j, (pieces, remaining) in enumerate(stock_used):
            data_to_export.append({
                t["Profile"]: profile,
                t["Stock Length"]: stock_length,
                t["Cut Index"]: j+1,
                t["Pieces Cut"]: pieces,
                t["Total Length Used"]: sum(pieces),
                t["Remaining Length"]: remaining
            })
    
    df = pd.DataFrame(data_to_export)
    
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=t["Cutting Plan"], index=False)
    
    workbook = writer.book
    worksheet = writer.sheets[t["Cutting Plan"]]
    worksheet.insert_image('H2', image_path)
    writer._save()

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

def main(data, settings_df, output_path, image_path, default_length, weight_error=12, steel_price=0):
    """Main function to run the optimization"""
    try:
        # If data is already a DataFrame, use it directly
        if isinstance(data, pd.DataFrame):
            data_cleaned = data
        else:
            # If it's a file path, load and clean the data
            data = load_data(data)
            data_cleaned = clean_data(data)
        
        # Run optimization
        results = optimize_cutting(data_cleaned, settings_df, default_length)
        
        # Generate outputs
        if results:  # Only generate outputs if we have results
            draw_cutting_plan(results, image_path)
            export_to_excel(results, output_path, image_path)
            
            # Calculate and return waste statistics
            waste_stats = calculate_waste_percentage(results)
            
            # Calculate weight statistics
            weight_stats = calculate_weight_stats(data_cleaned)
            
            # Calculate total weight and price
            total_weight = sum(weight_stats.values())
            adjusted_weight = total_weight + (total_weight * (weight_error/100))
            total_price = adjusted_weight * steel_price
            
            # Add weight and price info to stats
            stats = {
                'waste': waste_stats,
                'weight': {
                    'profiles': {k: round(v, 3) for k, v in weight_stats.items()},
                    'total': round(total_weight, 3),
                    'adjusted': round(adjusted_weight, 3),
                    'price': round(total_price, 3)
                }
            }
            
            return stats
        
    except Exception as e:
        raise Exception(f"Error in optimization process: {str(e)}")
