# Silo Design 
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from google.colab import files

# Output file
out_path = "silo_design_auto.xlsx"

wb = Workbook()

# -------- Inputs Sheet --------
ws_in = wb.active
ws_in.title = "Inputs"

inputs = [
    ("Diameter (D)", 4.0, "m"),
    ("Cylindrical height (h_c)", 6.0, "m"),
    ("Hopper height (h_h)", 2.0, "m"),
    ("Bulk density (ρ)", 800.0, "kg/m³"),
    ("Gravity (g)", 9.81, "m/s²"),
    ("Wall friction coeff (μ)", 0.3, "-"),
    ("Lateral pressure coeff (K)", 0.4, "-"),
    ("Wall thickness (t)", 0.01, "m"),
    ("Depth (z)", 8.0, "m"),
]

ws_in.append(["Parameter", "Value", "Units"])
for name, val, unit in inputs:
    ws_in.append([name, val, unit])

for col in range(1, 4):
    ws_in.column_dimensions[get_column_letter(col)].width = 25
for cell in ws_in[1]:
    cell.font = Font(bold=True)


# -------- Equations Sheet --------
ws_eq = wb.create_sheet("Equations")
ws_eq.append(["Calculation", "Equation"])
equations = [
    ("Radius", "r = D / 2"),
    ("Cross sectional area", "A = π r²"),
    ("Volume cylinder", "V_c = π r² h_c"),
    ("Volume hopper", "V_h = (1/3) π r² h_h"),
    ("Total Volume", "V = V_c + V_h"),
    ("Stored mass", "W = V × ρ"),
    ("Hydraulic radius", "R = D / 2"),
    ("Vertical stress (Janssen)", "σv(z) = (ρ g / K) × (1 - e^(-K z / (R μ)))"),
    ("Lateral pressure", "σh = K × σv"),
    ("Hoop stress", "f_h = σh × D / (2 t)"),
    ("Longitudinal stress", "f_l = σv × r / (2 t)"),
    ("Bearing pressure", "q = W / A"),
]
for name, eq in equations:
    ws_eq.append([name, eq])

ws_eq.column_dimensions["A"].width = 35
ws_eq.column_dimensions["B"].width = 80
for cell in ws_eq[1]:
    cell.font = Font(bold=True)


# -------- Calculations Sheet --------
ws_calc = wb.create_sheet("Calculations")
ws_calc.append(["Calculation", "Excel Formula", "Units"])

calc_formulas = [
    ("Radius (r)", "=Inputs!B2/2", "m"),
    ("Area (A)", "=PI()*(Inputs!B2/2)^2", "m²"),
    ("Volume cylinder", "=PI()*(Inputs!B2/2)^2*Inputs!B3", "m³"),
    ("Volume hopper", "=(1/3)*PI()*(Inputs!B2/2)^2*Inputs!B4", "m³"),
    ("Total Volume", "=C3+C4", "m³"),
    ("Stored mass (W)", "=C5*Inputs!B5", "kg"),
    ("Hydraulic radius (R)", "=Inputs!B2/2", "m"),
    ("Vertical stress σv", "=(Inputs!B5*Inputs!B6/Inputs!B7)*(1-EXP(-Inputs!B7*Inputs!B9/(C7*Inputs!B6)))", "Pa"),
    ("Lateral pressure σh", "=Inputs!B7*C8", "Pa"),
    ("Hoop stress f_h", "=C9*Inputs!B2/(2*Inputs!B8)", "Pa"),
    ("Longitudinal stress f_l", "=C8*C1/(2*Inputs!B8)", "Pa"),
    ("Bearing pressure q", "=C6/C2", "Pa"),
]

for name, formula, unit in calc_formulas:
    ws_calc.append([name, formula, unit])

ws_calc.column_dimensions["A"].width = 40
ws_calc.column_dimensions["B"].width = 55
ws_calc.column_dimensions["C"].width = 15
for cell in ws_calc[1]:
    cell.font = Font(bold=True)

# Save and download
wb.save(out_path)
files.download(out_path)
