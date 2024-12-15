import pandas as pd
import matplotlib.pyplot as plt
import os
import xlsxwriter

def load_data(file_path):
    """Load data from Excel files"""
    try:
        data = pd.read_excel(file_path, engine='openpyxl')
        return data
    except Exception as e:
        raise Exception(f"Error loading file {file_path}: {str(e)}")

def clean_data(data):
    """Clean and prepare the data"""
    try:
        # Extract relevant columns and remove rows with NaN values
        data_cleaned = data[['Profil', 'Qté', 'Long.']].dropna(subset=['Profil', 'Qté', 'Long.'])
        data_cleaned = data_cleaned[~data_cleaned['Profil'].str.contains('Total', na=False)]
        
        # Convert to proper types
        data_cleaned['Qté'] = pd.to_numeric(data_cleaned['Qté'], errors='coerce').fillna(1).astype(int)
        data_cleaned['Long.'] = pd.to_numeric(data_cleaned['Long.'], errors='coerce').fillna(0).astype(int)
        
        return data_cleaned
    except Exception as e:
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

def main(data, settings_df, output_path, image_path, default_length):
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
            return calculate_waste_percentage(results)
        
    except Exception as e:
        raise Exception(f"Error in optimization process: {str(e)}")
