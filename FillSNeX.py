from openpyxl import Workbook, load_workbook, worksheet, styles
import datetime
import argparse
import math
import random

date_line = 13
ERC_project_line = 15
other_projects_line = 28
teaching_line = 29
admin_line = 30
holiday_color_cell = "B44"
halfday_color_cell = "B45"
month_cell = "I9"
half_day_optional_color = "FFFFCCFF"
half_day_hours = 5
months = {1:"Jan",2:"Feb", 3:"Mar", 4:"Apr",5:"May",6:"Jun", 7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov", 12:"Dec"}


def derive_teaching_dict(teaching_days_and_hours_list):
    teaching_days_and_hours_dict = dict()
    for pair in teaching_days_and_hours_list.replace('{','').replace('}','').split(','):
        splitted_pair = pair.split(':')
        teaching_days_and_hours_dict[int(splitted_pair[0])] = int(splitted_pair[1])

    return teaching_days_and_hours_dict

def fill_line(working_tab, row, relevant_columns, total, min_daily, max_daily):
    filled_cells = [0 for c in relevant_columns]
    for i,c in enumerate(relevant_columns):
        if min_daily[i] > max_daily[i]:
            continue
        if total > max_daily[i]:
            hours = random.randint(min_daily[i],max_daily[i])
            working_tab[c+str(row)] = hours
            filled_cells[i] = hours
            max_daily[i] -= hours
            total -= hours
        else:
            working_tab[c + str(row)] = total
            filled_cells[i] = total
            max_daily[i] -= total
            total = 0
            break

    if total > 0:
        for i,c in enumerate(relevant_columns):
            if filled_cells[i] < max_daily[i]:
                hours = min((max_daily[i] - filled_cells[i]), total - filled_cells[i])
                filled_cells[i] += hours
                working_tab[c+str(row)] = filled_cells[i]
                total -= hours
                max_daily[i] -= hours
            if total == 0:
                break

    return filled_cells, max_daily

def calculate_and_fill_teaching_days(working_tab,relevant_columns,teaching_days_and_hours):
    teaching_hours = 0
    teaching_days = teaching_days_and_hours.keys()
    columns = [0 for i in range(len(relevant_cells))]
    for i, c in enumerate(relevant_columns):
        first_month_date = working_tab[month_cell]
        current_date = datetime.datetime(first_month_date.value.year, first_month_date.value.month, working_tab[c + str(date_line)].value)
        week_day = (current_date.weekday()+2)%7
        if week_day in teaching_days:
            working_tab[c + str(teaching_line)] = teaching_days_and_hours[week_day]
            teaching_hours += teaching_days_and_hours[week_day]
            columns[i] = teaching_days_and_hours[week_day]

    return teaching_hours, columns


def calculate_total_working_hours(working_tab, personal_max_hours):
    '''
    caount all working hours this moth by the holidays indicated by the colors of the cells
    :param working_tab: tab instance
    :return: total working hours this month, and the list of working columns.
    here we assume nobody works on half days, although it is untrue...
    '''
    hours = 0
    columns = []
    relevant_daily_max_hours = []
    for i in range(0,32):
        current_day_hours = args.average_daily_hours
        current_max = personal_max_hours
        if i < 25:
            column = chr(ord('B')+i)
        else:
            column = "A"+chr(i-25+ord('A'))
        if working_tab[str(column)+str(date_line)].value == "Total":
            break
        first_month_date = working_tab[month_cell]
        current_date = datetime.datetime(first_month_date.value.year,first_month_date.value.month,i+1)
        week_day = current_date.weekday()
        if week_day in [4,5]:
            continue
        if working_tab[column+str(date_line)].fill.start_color.value == working_tab[holiday_color_cell].fill.start_color.value:
            continue
        if working_tab[column+str(date_line)].fill.start_color.value == working_tab[halfday_color_cell].fill.start_color.value or \
                working_tab[column+str(date_line)].fill.start_color.value == half_day_optional_color:
            current_day_hours = half_day_hours
            current_max = half_day_hours

        hours += current_day_hours
        columns.append(column)
        relevant_daily_max_hours.append(current_max)

    return hours, columns, relevant_daily_max_hours

def InitParser():
    parser = argparse.ArgumentParser(description='Welcome to the automatic SNex filler!')
    parser.add_argument('--document_path', type=str, required=True ,help='path to the excel sheet to be filled')
    parser.add_argument('--tab', type=int, default=9 ,help='The number of the required month, for example 1 for January, 5 for May')
    parser.add_argument('--months', nargs='+', type=int, default=[] ,help='list of month indices')
    parser.add_argument('--ERC_percentage', type=float, default=100 ,help='percentage of time working on the ERC projects')
    parser.add_argument('--average_daily_hours', type=float, default=8, help='The average hours working a day')
    parser.add_argument('--min_daily_hours', type=float, default=2, help='The minimal hours working a day')
    parser.add_argument('--max_daily_hours', type=float, default=9, help='The maximal hours working a day')
    parser.add_argument('--is_admin', type=lambda x: (str(x).lower() in ['true', '1', 'yes']), default=False,
                        help='True if you are have additional administrative works, ignoring all administrative values if false')
    parser.add_argument('--average_admin_daily_hours', type=float, default=1, help='The average administrative hours working a day')
    parser.add_argument('--min_admin_daily_hours', type=float, default=0, help='The minimal administrative hours working a day')
    parser.add_argument('--max_admin_daily_hours', type=float, default=4, help='The maximal administrative hours working a day')
    parser.add_argument('--teaching_days_and_hours', type=str, default={}, help='a dictionary with the teaching day '
                                                                                'number as the key (Sunday is 1) and '
                                                                                'the value is number of hours, '
                                                                                'the form is {1:2, 4:3}')
    return parser


if __name__ == "__main__":
    parser = InitParser()
    args = parser.parse_args()
    workbook = load_workbook(args.document_path)
    if len(args.months) == 0:
        months_n = [args.tab]
    else:
        months_n = args.months

    for month_n in months_n:
        working_tab = workbook.get_sheet_by_name(months[month_n])
        monthly_hours, relevant_cells, daily_max_hours = calculate_total_working_hours(working_tab, args.max_daily_hours)

        monthly_teaching_hours = 0
        max_admin_daily_hours = 0
        filled_teaching = [0 for i in range(len(relevant_cells))]
        if args.is_admin:
            print("Hello manager! It is now time to fill in the administrative working hours.")
            print("Notice that we assume that teaching happens even in half days and there is no such thing as \"MATKONET\"")
            try:
                teaching_days_and_hours_dict = derive_teaching_dict(args.teaching_days_and_hours)
                monthly_teaching_hours, filled_teaching = calculate_and_fill_teaching_days(working_tab,relevant_cells,teaching_days_and_hours_dict)
                daily_max_hours = [daily_max_hours[i]-filled_teaching[i] for i in range(len(filled_teaching))]
            except:
                print("Could not fill in teaching hours, check the format of your supplied dictionary, "
                      "it should look like- {1:2,4:1} for teaching 2 hours on Sunday and one hour on Wednesday")
            monthly_admin_hours = args.average_admin_daily_hours*len(relevant_cells)
            max_admin_daily_hours = args.max_admin_daily_hours
        else:
            monthly_admin_hours = 0
            monthly_teaching_hours = 0

        monthly_ERC_hours = math.ceil(monthly_hours * args.ERC_percentage / 100.0)
        monthly_other_hours = monthly_hours - monthly_ERC_hours - monthly_admin_hours - monthly_teaching_hours

        print("calculated total hours, ERC hours and other project this month:")
        print(monthly_hours, monthly_ERC_hours, monthly_other_hours)

        filled_ERC, daily_max_hours = fill_line(working_tab,ERC_project_line,relevant_cells,monthly_ERC_hours,
                               [0 for i in range(len(relevant_cells))], daily_max_hours)

        filled_others, daily_max_hours = fill_line(working_tab,other_projects_line,relevant_cells,monthly_other_hours,
                  [max(0,args.min_daily_hours - filled_ERC[i] - filled_teaching[i]- max_admin_daily_hours) for i in range(len(filled_ERC))], daily_max_hours)

        if args.is_admin:
            fill_line(working_tab,admin_line, relevant_cells, monthly_admin_hours,
                  [args.min_admin_daily_hours for i in range(len(relevant_cells))],
                  [min(args.max_admin_daily_hours, daily_max_hours[i]) for i in range(len(relevant_cells))])

    workbook.save(args.document_path.replace('.xlsx','_filled.xlsx'))

    print("filled form saved as ", args.document_path.replace('.xlsx','_filled.xlsx'))

