import io
import xlsxwriter
from fastapi import FastAPI, Response

app = FastAPI()

@app.get("/prepare_excel_data")
async def prepare_excel_data(response: Response):
    headers = [
        'RN 10HL', 'RN 20HL', 'Royals Next Total', 'RG 20HL', 'RG 10HL', 'Royals Gold Total', 'RLS 10HL', 'Royals LC 10HL', 
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
