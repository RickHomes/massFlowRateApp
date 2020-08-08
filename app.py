from flask import Flask
from flask import render_template, request

from datetime import datetime
import math
import os
import os.path
from os import system
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
app = Flask(__name__)

# ================================================================
# Global variables and constants
g = 9.81  # g(m/s)
rho = 760  # Grain bulk density(kg/m~3
k = 1.4  # k is the shape factor
pi = 3.142  # pi
d_geo = 0.0085  # d is the geometric mean diammeter of the grain
c_fr = 0.75  # Friction/discharge coefficient
# models = ['Beverloo', 'BCP', 'Tudor', 'EXIT']
html_output = []
# MAIN FUNCTION
# ================================================================


@app.route('/', methods=['GET', 'POST'])
def main():
    #  remove files
    def clear_file():
        os.remove('static/Model_reports.xlsx')
        # main()
    # Date today
    date_object = datetime.now()
    # to create a new file if none exist
    if os.path.isfile('static/Model_reports.xlsx') == False:
        filepath = 'static/Model_reports.xlsx'
        wb = openpyxl.Workbook()
        wb.save(filepath)
    # use existing file
    wb = openpyxl.load_workbook('static/Model_reports.xlsx')
    ws = wb.active

    if request.method == 'POST':
        model = request.form['models']
        mode = request.form['modes']
        first_value = request.form['first_value']
        second_value = request.form['second_value']
        model_steps = request.form['model_steps']
        save_mode = request.form['save_modes']

        if int(save_mode) == 2:
            clear_file()
        else:
            pass

        # computation
        if model == "beverloo":
            # print("beverloo mode activated")
            # print(mode)
            html_output.clear()

            def beverloo_model():
                # print(mode)
                if mode == "manual":
                    print('Manual Mode\n\n Enter the diammeter of the orifice (mm)')
                    model1_diammeter = int(input())
                    area = (pow((model1_diammeter/1000), 2)*pi)/4
                    de = (model1_diammeter/1000)-(k*d_geo)  # in m
                    q = area*math.sqrt(de*g)*3.6
                    mass_f = (q*rho*c_fr)
                    print(
                        f'\t\tGrain Mass Flow Rate for {model1_diammeter}mm == {mass_f} Kg/h')
                    main()
                elif mode == "automatic":
                    print("automatic mode activated")
                    if '' not in [first_value, second_value, model_steps]:

                        start = float(first_value)
                        stop = float(second_value)
                        step = float(model_steps)
                        result = []

                        def frange(start, stop, step):
                            while start < stop:
                                yield start
                                start += step

                        for r in frange(start, stop, step):

                            data = r
                            area = (pow((r/1000), 2)*pi)/4
                            de = r-(k*d_geo)  # in m
                            q = area*math.sqrt(de*g)
                            mass_f = q*rho*3600
                            # print(
                            #     f'Grain Mass Flow for {r}mm == {mass_f} Kg/h discharge {q} \n\n a copy of the results has been saved')

                            entry = [r, mass_f]
                            entry_2 = {'diameter': r, 'flow': mass_f}
                            str(entry).replace("(", "").replace(")", "")
                            str(entry_2).replace("(", "").replace(")", "")

                            result.append(entry)
                            html_output.append(entry_2)
                        # print(html_output)
                        # print(result)
                        # add column headings. NB. these must be strings

                        ws.append(
                            ['Diammeter(mm)', 'Mass Flow(Kg/h)'])

                        for row in result:
                            ws.append(row)
                        tab = Table(displayName="Table" +
                                    str({date_object.microsecond}), ref="A1:E5")

                        # Add a default style with striped rows and banded columns
                        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
                        tab.tableStyleInfo = style

                        '''
                            Table must be added using ws.add_table() method to avoid duplicate names.
                            Using this method ensures table name is unque through out defined names and all other table name.
                            '''
                        ws.add_table(tab)
                        wb.save("static/Model_reports.xlsx")
            beverloo_model()
    return render_template('index.html', title='calculate', heading="MASS FLOW RATE SIMULATION MODELS", result=html_output)

    main()
