"""Extract contact list from screenshots and save to Excel."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

contacts = [
    # R section
    ("James Rasque", "IT Business Analyst Supervisor", "Kwik Trip Inc"),
    ("Abhishek Rastogi", "Lead Specialist, Business Cont...", "National Grid"),
    ("Jeannine Rau", "SCSM", "SAP AMERICA, INC."),
    ("Koushika Ravikumar", "Business Analyst", "KATBOTZ LLC"),
    ("Prachi Ray", "Marketing - SAP", "Tech Mahindra Limited"),
    ("Vasee Rayan", "Senior Vic...", "SAP Labs, LLC. (Newtown Square)"),
    ("Carlos Recio", "Vice President Testing Practice", "Innovise IT"),
    ("Araja Reddy", "Managing Director", "Deloitte Consulting"),
    ("Sasi Reddy", "Sr. SAP Applicati...", "Sun Chemical Corporation"),
    ("Nico Reichen", "", "PricewaterhouseCoopers GmbH Wirtschaftsp..."),
    ("Caitlin Reidy", "", "Icertis Inc."),
    ("Kristen Reinken", "Journey & Change Management L...", "Accenture"),
    ("Efrain Reyes", "SAP Function...", "Aecon Construction Group Inc."),
    ("Luis Reyes", "Regional Sales Director SSLATAM", "Vistex Inc."),
    ("Dan Rhoadhouse", "Process Ar...", "The Scotts Miracle-Gro Company"),
    ("Fernando Riccardi", "Chief Operating Officer", "Exed Consulting"),
    ("Devin Rich", "Chief Info. Off.", "Lansing Building Products, LLC"),
    ("Trey Riley", "Senior Director", "Leprino Foods Company"),
    ("Lisa Rinaldi", "Finance Lead, ED...", "General Electric Company"),
    ("Angela Rios", "CIO", "Centro de Sistemas y Negocios S.A."),
    ("Jerome Robertson", "Director, AI Driven Business Insights", "Pfizer"),
    ("Lauren Robinson", "", "Sparks"),
    ("Devin Robnick", "Head of NA Field Market...", "SAP AMERICA, INC."),
    ("Virginia Roda", "Chief Technology Off", "Southbay S.R.L."),
    ("Alexander Rodde", "Head of Corporate Portfolio", "SAP SE"),
    ("Sebastian Rode", "Global ERP Delivery Lead", "Roche"),
    ("Hernan Rodriguez", "Practice Manager", "Nearshore Argentina S.A."),
    ("Josiane Roger", "", "CGI Information Systems and Management C..."),
    ("Laurie Rogosheske", "Director E...", "Wolters Kluwer United States Inc."),
    ("Rachel Romanoski", "Director - SAP SCM...", "SAP America (Houston)"),
    ("Diana Romaya", "SAP...", "Deloitte Touche Tohmatsu Services, LLC"),
    ("John Rooney", "Regional Partner Director", "Stonebranch"),
    ("Susanne Ross", "Senior Partner Marketin...", "Coveo Solutions Inc."),
    ("Cori Rothgery", "Data & A...", "DURACELL U.S. OPERATIONS, INC."),
    ("Sylvain Roy", "It Director", "Heroux-Devtek"),
    ("Renee Rubino", "Senior Manager of Alliance Partnershi...", "SMC3"),
    ("Felix Rubio", "Industry Advisor Latin America", "SAP Colombia"),
    ("Frank Ruetten", "Client Partner", "IBM Deutschland GmbH"),
    ("Frank Ruland", "", "PricewaterhouseCoopers GmbH Wirtschaftsp..."),
    ("Roxanne Ryan", "Director, Executive Progr...", "SAP AMERICA, INC."),
    ("Scott Ryerson", "Partner Executive", "Capgemini America, Inc."),
    ("Royce Ryu", "Head of Services Deliver...", "SAP Japan Co., Ltd."),
    ("Natalie Röthel", "Executive Assistant", "McKinsey & Company"),
    ("mariano rivarola", "Gerente Industrial Regional", "HZ GROUP S.R.L."),
    # S section
    ("Hanumantha Rao Sabbireddy", "New Business", "Tech Mahindra"),
    ("Andrew Sackett", "Director of Partnerships", "Collibra Belgium BV"),
    ("Rinchen Sahni", "Digital Marketing Manager", "KATBOTZ LLC"),
    ("KG Saito", "Vice President", "MACNICA, INC."),
    ("Makoto Saito", "IT Development Dept", "Tokyo Electron Limited"),
    ("Tito Salazar", "Partner Growth Manager", "OneRail"),
    ("Sarah Salzman", "Regional VP of Sales", "Basware, Inc."),
    ("Julissa Samuels", "Gerente de aplica...", "American Sportswear, S.A."),
    ("Brent Sandberg", "Project Manager, Business Transf...", "Carmeuse"),
    ("Sachin Sangle", "VP", "ImpactQA"),
    ("Marco Sanna", "", "cbs Corporate Business Solutions Unternehm..."),
    ("KATSUYUKI SHINZEKI", "Assistant Manager", "Marubeni Corporation"),
    ("JASBIR SINGH", "Associate Partner, Global...", "IBM India Pvt Ltd"),
]

wb = Workbook()
ws = wb.active
ws.title = "Contacts"

headers = ["Name", "Role", "Company"]
ws.append(headers)

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="left", vertical="center")

for name, role, company in contacts:
    ws.append([name, role, company])

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 45
ws.column_dimensions["C"].width = 50

ws.freeze_panes = "A2"

output = "/home/user/test-one/contacts.xlsx"
wb.save(output)
print(f"Saved {len(contacts)} contacts to {output}")
