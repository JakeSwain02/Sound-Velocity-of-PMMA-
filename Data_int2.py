import pprint
import xlrd
import statistics as stat
import matplotlib.pyplot as plt


def get_data(file, sheet):
    # set of commands that imports the data from an exel file
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_name(sheet)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    return data

def create_dataset(data):
    # this selects the spefic data used which is time and signal.
    cords, x_all, y_all = [], [], []

    for row in range((len(data) -1)):
        if data[row][1] >= 0:
            if data[row + 1][1] - data[row][1] < 1.5:
                # the x-values are the time which is the second collunm (1st in terms of index value) for each row
                x = round(data[row][1])
                # the y-value is signal which is the maginitude of the 3rd and 4th columns for each row
                y = ((data[row][2] ** 2) + (data[row][3] ** 2)) ** .5
                # the values are both stored together are cordnates for the calculation and seperate for the graphing
                cords.append([x,y])
                x_all.append(x)
                y_all.append(y)
    return cords, x_all, y_all

def find_max(cords):
    # this function finds all of the peak values in the graph
    max_points, x_max, y_max = [], [], []
    for cord in range((len(cords) -2)):
        # finds the maxiumum by only accpeting points where the y-value of the point behind and infront are less
        if cords[cord][1] < cords[cord + 1][1]:
            if cords[cord + 1][1] > cords[cord + 2][1]:
                x = cords[cord + 1][0]
                y = cords[cord + 1][1]
                max_points.append([x , y])
                x_max.append(x)
                y_max.append(y)
    return max_points, x_max, y_max

def find_min(cords):
    # same process for max points, but for min points
    min_points, x_min, y_min = [], [], []
    for cord in range((len(cords) -2)):
        if cords[cord][1] > cords[cord + 1][1]:
            if cords[cord + 1][1] < cords[cord + 2][1]:
                x = cords[cord + 1][0]
                y = cords[cord + 1][1]
                min_points.append([x , y])
                x_min.append(x)
                y_min.append(y)
    return min_points, x_min, y_min

def find_distance(max_of_max_cords):
    distances, distances_above_avg, distance_first_ten, distances_clean = [], [], [], []
    # finds the distance of the max points by the x-value differnce of the point and the one infront of it.
    for point in range((len(max_of_max_cords) -1)):
        distance = max_of_max_cords[point + 1][0] - max_of_max_cords[point][0]
        distances.append(distance)
    avg_distance = stat.mean(distances)

    # elinemates invalid data by removing data that is below the average, after the 10th values, and outliers
    # this is becuase if the data is below the average value a maxiumm was split in half, and after 10 values data was inconsistant
    for distance in distances:
        if distance > avg_distance:
            distances_above_avg.append(distance)
    for i in range(11):
        distance_first_ten.append(distances_above_avg[i])
    mean = stat.mean(distance_first_ten)
    for distance in distance_first_ten:
        if (mean * .8) < distance < (mean * 1.2):
            distances_clean.append(distance)
    avg_distance = stat.mean(distances_clean)

    return avg_distance

def find_distance_plan_b(max_of_max_cords, min_of_min_cords):
    # this is a backup incase the first on doesnt work. some of the weaker datasets cause that erros and use this instead
    max_dist = abs(max_of_max_cords[0][0] - max_of_max_cords[1][0])
    min_dist = abs(min_of_min_cords[0][0] - min_of_min_cords[1][0])

    avg_distance = (max_dist + min_dist) / 2

    return avg_distance

def find_velocity(avg_distance, gap_distance):

    # formula that solve of velocity
    velocity = (2 * gap_distance) / (avg_distance * (10**-12))

    return velocity

def find_points_between(cords):
    # this is made for finding the beat freq. It creates two lines, a line tracing the min and a line tracing the max
    # this is so the program can find a beat frequcny when the min and the max line approach eachother
    x_points, y_points= [], []
    for cord in range(len(cords) -1):
        if cords[cord + 1][0] != cords[cord][0]:
            # uses y - mx + b, inorder to interplate the lines inbetween the max and min points
            m = (((cords[cord + 1][1]) - (cords[cord][1])) / ((cords[cord + 1][0]) - (cords[cord][0])))
            b = cords[cord][1] - (m * cords[cord][0])
            for x in range(cords[cord][0], cords[cord + 1][0]):
                y = (m * x) + b
                x_points.append(x)
                y_points.append(y)
    return x_points, y_points

def find_beat_freq(max_of_max_cords, min_of_min_cords):
    x_max_of_max_all, y_max_of_max_all, x_min_of_min_all, y_min_of_min_all, mins, count, mins_final  = [],[],[],[], [], -1, []

    x_max_of_max_all, y_max_of_max_all = find_points_between(max_of_max_cords)
    x_min_of_min_all, y_min_of_min_all = find_points_between(min_of_min_cords)

    # balences out len of the min and max data set so they can be accessed with a common index value
    len_diff = abs(x_max_of_max_all[0] - x_min_of_min_all[0])
    if x_max_of_max_all[0] >= x_min_of_min_all[0]:
        for value in range(len_diff):
            x_min_of_min_all.remove(x_min_of_min_all[0])
            y_min_of_min_all.remove(y_min_of_min_all[0])
    else:
        for value in range(len_diff):
            x_max_of_max_all.remove(x_max_of_max_all[0])
            y_max_of_max_all.remove(y_max_of_max_all[0])

    if len(x_max_of_max_all) >= len(x_min_of_min_all):
        shortest = len(x_min_of_min_all)
    else:
        shortest = len(x_max_of_max_all)



    # finds the fist four mininum distance between the mins and the maxes
    for i in range(shortest - 2):
        if (y_max_of_max_all[i] - y_min_of_min_all[i]) > (y_max_of_max_all[i + 1] - y_min_of_min_all[i + 1]):
            if (y_max_of_max_all[i + 1] - y_min_of_min_all[i + 1]) < (y_max_of_max_all[i + 2] - y_min_of_min_all[i + 2]):
                if count < 3:
                    mins.append(x_max_of_max_all[i + 1])
                    count = count + 1

    # makes both min and max start at an x value of 0
    for i in range(x_max_of_max_all[0]):
        y_max_of_max_all.insert(0, 0)
        y_min_of_min_all.insert(0, 0)

    # finds the correct values by comparing simular potential values and elimnating the worse one
    i = 0
    while i + 1 < len(mins):
        if abs(mins[i] - mins[i + 1]) < 100:
            current_distance = abs(y_max_of_max_all[(mins[i])] - y_min_of_min_all[(mins[i])])
            next_distance = abs(y_max_of_max_all[(mins[i + 1])] - y_min_of_min_all[(mins[i + 1])])
            if current_distance > next_distance:
                mins.remove(mins[i])
        i = i + 1

    # takes out leftover values if neccisary
    while len(mins) > 2:
        mins.remove(mins[-1])

    beat_distance = abs(mins[1] - mins[0])

    return beat_distance, mins

def display_trail(sheet, avg_distance, Al_velocity, beat_distance, PMMA_velocity):
    # displays information for each trial
    print('For', sheet)
    print('Crest Distance', avg_distance)
    print('Aluminium Velocity', round(Al_velocity, 3), 'm/s')
    if beat_distance != 'None':
        print('Beat Distance =', beat_distance)
        print('PMMA Velocity', round(PMMA_velocity, 3), 'm/s')
    print('...')

def graph(x_all,y_all, x_max_of_max, y_max_of_max, x_min_of_min, y_min_of_min, x_beat_cords, sheet):
    # graphs data using matplotlib
    plt.xlabel('Time')
    plt.ylabel('Signal')
    plt.title(sheet)
    plt.xlim(min(x_all),max(x_all))
    plt.ylim(min(y_all), max(y_all))

    plt.plot(x_all, y_all, label = 'All')
    plt.plot(x_max_of_max, y_max_of_max, color = 'r', label = 'Max and Min')
    plt.plot(x_min_of_min, y_min_of_min,  color = 'r')

    if x_beat_cords != 'None':
        for cord in range(len(x_beat_cords)):
            plt.axvline(x = x_beat_cords[cord], color = 'c', label = 'Beat Frequency')

    plt.plot()
    plt.legend()
    plt.show()

def find_results(velocity_data, Al_expected, PMMA_expected, trial):
    # extracts data
    Al, PMMA = [], []
    for row in range(len(velocity_data)):
        Al.append(velocity_data[row][0])
        if velocity_data[row][1] != 'None':
            PMMA.append(velocity_data[row][1])
    # sets mean and IQR
    Al_mean_org, PMMA_mean_org = stat.mean(Al), stat.mean(PMMA)
    Al_IQR, PMMA_IQR = (.25 * stat.mean(Al)) , (.25 * stat.mean(PMMA))

    # removes outliers by IQR
    for Al_val in Al:
        if abs(Al_val - Al_mean_org) > Al_IQR:
            Al.remove(Al_val)
    for PMMA_val in PMMA:
        if abs(PMMA_val - PMMA_mean_org) > PMMA_IQR:
            PMMA.remove(PMMA_val)
    Al_mean, PMMA_mean = stat.mean(Al), stat.mean(PMMA)
    #  finds error
    if Al_expected == 'None':
        Al_error = 'None'
    else:
        Al_error = (abs(Al_mean - Al_expected) / Al_expected) * 100

    if PMMA_expected == 'None':
        PMMA_error = 'None'
    else:
        PMMA_error = (abs(PMMA_mean - PMMA_expected) / PMMA_expected) * 100

    # distplays information
    print('For', trial)
    print('Velocity of Sound in Aluminium:', round(Al_mean,2))
    if Al_error != 'None':
        print('With', round(Al_error,2) ,'% Uncertainty')
    if len(Al) <= 2:
        print('Warning: Low Amount of Sufficient Data for Aluminium, Results May be Unreliable')
    print('Velocity of Sound in PMMA:', round(PMMA_mean,2))
    if PMMA_error != 'None':
        print('With', round(PMMA_error,2),'% Uncertainty')
    if len(PMMA) <= 2:
        print('Warning: Low Amount of Sufficient Data for PMMA, Results May be Unreliable')
    print('...')

    Results = [Al_mean, Al_error, PMMA_mean, PMMA_error]

    return Results

def run_sheets(file, sheet, gap_distance_Al, gap_distance_PMMA, set):
    try:
        # formats data
        data = get_data(file, sheet)
        all_cords, x_all, y_all = create_dataset(data)
        # finds mins and maxes
        min_cords, x_min, y_min = find_min(all_cords)
        min_of_min_cords, x_min_of_min, y_min_of_min, = find_min(min_cords)
        max_cords, x_max, y_max = find_max(all_cords)
        max_of_max_cords, x_max_of_max, y_max_of_max, = find_max(max_cords)
        # finds beat frequency
        beat_distance, x_beat_cords = find_beat_freq(max_of_max_cords, min_of_min_cords)
        avg_distance = find_distance(max_of_max_cords)
        # finds velocitys
        Al_velocity = find_velocity(avg_distance, gap_distance_Al)
        PMMA_velocity = find_velocity(beat_distance, gap_distance_PMMA)
        # displays and graphs information
        display_trail(sheet, avg_distance, Al_velocity, beat_distance, PMMA_velocity)
        graph(x_all,y_all, x_max_of_max, y_max_of_max, x_min_of_min, y_min_of_min, x_beat_cords, sheet)
        # holds data
        set.append([Al_velocity, PMMA_velocity])
    except:
        # does simple method for weak data
        try:
            # finds mins and maxes
            min_cords, x_min, y_min = find_min(all_cords)
            min_of_min_cords, x_min_of_min, y_min_of_min, = find_min(min_cords)
            max_cords, x_max, y_max = find_max(all_cords)
            max_of_max_cords, x_max_of_max, y_max_of_max, = find_max(max_cords)
            avg_distance = find_distance_plan_b(max_of_max_cords, min_of_min_cords)

            Al_velocity = find_velocity(avg_distance, gap_distance_Al)

            # displays and graphs information
            display_trail(sheet, avg_distance, Al_velocity, 'None', 'None')
            graph(x_all,y_all, x_max_of_max, y_max_of_max, x_min_of_min, y_min_of_min, 'None', sheet)
            # holds data
            set.append([Al_velocity, 'None'])

        except:
            # if all else fails
            print('For',sheet, 'Data insufficient')
            print('...')

def final_ans():
    # where the user interacts to decide what sheets to run
    # formateds as run_sheets(filepath, sheet_man, Al_width, PMMA_width, group)
    Control, Temp_70, Temp_100, Temp_70_33, Temp_70_66 = [],[],[],[], []
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C1',  (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C2', (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C3',  (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C4', (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C5',  (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C6', (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C7',  (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', 'C8',  (80 * (10 ** -9)), (158 * (10 ** -9)), Control)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT1', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT2', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT3', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT4', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT5', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70CT6', (80 * (10 ** -9)), (217 * (10 ** -9)), Temp_70)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT1', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT2', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT3', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT4', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT5', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '100CT6', (80 * (10 ** -9)), (210 * (10 ** -9)), Temp_100)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T1', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T2', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T3', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T4', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T5', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C33T6', (80 * (10 ** -9)), (167 * (10 ** -9)), Temp_70_33)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T1', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T2', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T3', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T4', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T5', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)
    run_sheets(r'/users/2020jswain/documents/Foley_Data.xlsx', '70C66T6', (80 * (10 ** -9)), (189 * (10 ** -9)), Temp_70_66)

    # finds fininal results
    # find_results(group, expected Al velocity , expected PMMA velocity, 'displayed name of group')
    # if there is no expected value put 'None'
    Results_Control = find_results(Control, 6320, 2270, 'Control (23 Degrees)' )
    Results_Temp_70 = find_results(Temp_70, 6320, 'None', '70 Degrees')
    Results_Temp_70_33 = find_results(Temp_70_33, 6320, 'None', '70 Degrees, 33%')
    Results_Temp_70_66 = find_results(Temp_70_66, 6320, 'None', '70 Degrees, 66%')
    Results_Temp_100 = find_results(Temp_100, 6320, 'None', '100 Degrees')


# calls entire program
final_ans()
