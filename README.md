# CV-Cyclic-Voltammetry
Interactive desktop app for cyclic voltammetry analysis with smoothing, baseline selection, peak &amp; derivative computation, and Excel export.

# CVision

**Analysis of the cyclic voltammogram**  
Version: 2.0.0

## Opis
CVision is a desktop application written in PyQt6 + pyqtgraph used to:
- CV data loading (E [mV], I_oxidation [μA], I_reduction [μA]),
- optional smoothing (Savitzky-Golay),
- interactive selection and manual editing of the baseline,
- calculation of peak parameters (x_peak, y_peak, height/depth, E₁/₂),
- analysis of the second derivative and detection of zero places,
- export the results to an Excel (.xlsx) file.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/StarGate3/cv-cyclic-voltammetry.git
   cd cv-cyclic-voltammetry

2. Create and activate a virtual environment:
    python -m venv venv
    source venv/bin/activate    # Linux/macOS
    venv\Scripts\activate       # Windows

3. Install dependencies:
    pip install -r requirements.txt

4. Launch the application:
    python main.py

## Using
1. Help – program manual.

2. About – version and author information.

3. Select the type of measurement (Oxidation/Reduction).

4. Load the data with the “Select data file” button.

5. Optionally smooth the data (do not increase the window >15).

6. Point to the baseline range, and then adjust manually.

7. Click “Calculate peak parameters”.

8. (For irreversible processes) “Calculate second derivative” → select range → “Find zero places”.

9. Export everything to Excel via “Export to Excel.”.

## Optional settings
1. Light/dark mode
2. Manual editing of axes (button “Edit axis settings”)

## License
MIT-licensed project.

## Author
StarGate3
GitHub: https://github.com/StarGate3
