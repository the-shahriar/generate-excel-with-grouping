import io
import xlsxwriter
from fastapi import FastAPI, Response

app = FastAPI()

@app.get("/prepare_excel_data")
async def prepare_excel_data(response: Response):
    headers = [
        'B&H BG 20HL', 'B&H BG PUG 12HL', 'B&H BG LEP 20HL', 'B&H BG LEP 12HL', 'B&H BG 12HL', 
        'B&H BG Total', 'B&H CRUSH 20HL', 'B&H Breeze 20HL LEPP', 'B&H Fusion 20HL LEPP', 
        'B&H SW LEP 20HL', 'B&H SW 20HL', 'B&H SW Total', 'B&H SF LEP 12HL', 'B&H SF LEP 20HL', 
        'B&H SF 20HL', 'B&H SF 12HL', 'B&H SF Total', 'Alchemy 7MG', 'Alchemy MIX1', 'Alchemy Total', 
        'B&H PT 20HL', 'B&H Family', 'JP SW 20HL', 'JPGL 20HL', 'JPGL 12HL', 'JPGL Total', 'JP SP 20HL', 
        'JP Family', 'CAP 20HL', 'Capstan Family', 'SRF 20HL', 'SRF 10HL', 'SRFT Total', 'Star Switch 20HL', 
        'SRFT Family', 'PL 20HL', 'Pilot Family', 'HWD 20HL', 'HWG 20HL', 'Hollywood Total', 'Hollywood Family', 
        'Derby Style 10HL', 'Derby Style 20HL', 'Derby ST 10HL LEP', 'Derby ST 20HL LEP', 'Derby Style Total', 
        'DB Deluxe 10HL', 'DB Deluxe 20HL', 'DB Deluxe Total', 'DB Select 10HL', 'DB Select 20HL', 'DB Select Total', 
        'Derby 10HL', 'Derby 20HL', 'Derby 20HL LEP', 'Derby 10HL LEP', 'Derby Total', 'Derby Family', 'RN 10HL', 
        'RN 20HL', 'Royals Next Total', 'RG 20HL', 'RG 10HL', 'Royals Gold Total', 'RLS 10HL', 'Royals LC 10HL', 
        'Royals LC 20HL', 'Royals LC Total', 'Royals Family', 'LS OG 20HL', 'LS RED 20HL LEPP', 'LS FT 20HL LEPP', 
        'LS BC 20HL LEPP', 'LS CC 20HL', 'Lucky Family'
    ]

    buffer = await generate_excel(headers)

    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    response.headers["Content-Disposition"] = "attachment; filename=report_292947673646634.xlsx"
    response.headers["Content-Length"] = str(len(buffer))

    return Response(content=buffer, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

async def generate_excel(headers: list) -> bytes:
    # Create a new workbook and add a worksheet
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('worksheet')

    # Create cell formats
    common_format = workbook.add_format({
        'font_size': 10,
        'border': 1
    })

    row = 0
    col = 0
    col_add = 0

    for element in headers:
        # insert family
        if ' Family' in element:
            worksheet.write(row, col + col_add, element, common_format)
            worksheet.set_column(col + col_add, col + col_add, 20)
            worksheet.set_column(col + col_add, col + col_add, 20, common_format, { 'level': 1,'hidden': True,'collapsed':True })
            col_add += 1
        # insert brand
        elif ' Total' in element:
            worksheet.write(row, col + col_add, element, common_format)
            worksheet.set_column(col + col_add, col + col_add, 20)
            worksheet.set_column(col + col_add, col + col_add, 20, common_format, { 'level': 2,'hidden': True,'collapsed':True })
            col_add += 1
        # general data
        else:
            worksheet.write(row, col + col_add, element, common_format)
            worksheet.set_column(col + col_add, col + col_add, 20)
            worksheet.set_column(col + col_add, col + col_add, 20, common_format, { 'level': 3,'hidden': True,'collapsed':True })
            col_add += 1

    # Insert total and volume manually
    worksheet.write(row, col + col_add, "Low Segment", common_format)
    worksheet.set_column(col + col_add, col + col_add, 20)
    worksheet.set_column(col + col_add, col + col_add, 20, common_format, { 'level': 1,'hidden': True,'collapsed':True })
    col_add += 1

    worksheet.write(row, col + col_add, "Total", common_format)
    worksheet.set_column(col + col_add, col + col_add, 20)
    col_add += 1

    worksheet.write(row, col + col_add, "Volume BY", common_format)
    worksheet.set_column(col + col_add, col + col_add, 20)

    # Close the workbook and write the output to a buffer
    workbook.close()
    output.seek(0)

    return output.read()
