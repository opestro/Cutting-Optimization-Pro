
# 🔧 Cutting Optimization Tool 🛠️

This is a Python tool to optimize the cutting of profiles based on given stock lengths. It uses the provided settings and working files to calculate the optimized cutting plans and visualize the results.

## 📦 Installation / Installation / التثبيت

1. **Clone the repository / Cloner le dépôt / استنساخ المستودع:**
   ```bash
   git clone <repository_url>
   cd <repository_directory>
   ```

2. **Create a virtual environment / Créer un environnement virtuel / إنشاء بيئة افتراضية:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install the required packages / Installer les packages requis / تثبيت الحزم المطلوبة:**
   ```bash
   pip install -r requirements.txt
   ```

## 📚 Required Packages / Packages Requis / الحزم المطلوبة

- pandas
- matplotlib
- xlsxwriter
- xlrd
- openpyxl

## 🛠️ Usage / Utilisation / الاستخدام

1. **Prepare the working file and settings file / Préparez le fichier de travail et le fichier de paramètres / قم بإعداد ملف العمل وملف الإعدادات:**
   - Ensure the working file (`A_liste appro[2].xlsx`) contains the profiles and their quantities.
   - Ensure the settings file (`settings.ods`) contains the default stock lengths for profiles.

2. **Run the optimization / Exécutez l'optimisation / قم بتشغيل التحسين:**
   ```bash
   python main.py
   ```

3. **Check the output / Vérifiez la sortie / تحقق من المخرجات:**
   - The optimized cutting plan will be saved as an Excel file (`optimized_cutting_plan_xlsx.xlsx`).
   - The cutting plan visualization will be saved as an image (`cutting_plan_xlsx.png`).

## 📁 Files / Fichiers / الملفات

- `main.py`: Main script to run the optimization.
- `cuttingOptimization.py`: Contains the functions for data loading, cleaning, optimization, and visualization.
- `requirements.txt`: List of required packages.
- `A_liste appro[2].xlsx`: Example working file.
- `settings.ods`: Example settings file.

## 📝 Example / Exemple / مثال

Here is an example command to run the optimization:
```bash
python main.py
```

## 📜 License / Licence / الترخيص

This project is licensed under the MIT License.
