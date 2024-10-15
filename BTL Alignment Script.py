import FreeCAD as App ; import Draft
import FreeCADGui as Gui ; import Part
from scipy.optimize import minimize
from tkinter import *
import pandas as pd
import numpy as np 
import math 

# Read the excel file
csv_file_path = "csv file path/BTL Alignment Spreadsheet.csv" #<- insert file path
df = pd.read_csv(csv_file_path)

# Open the visualizer
doc = App.newDocument()

# Plot co-ordinate data for the circumfrance
def create_points(coords):
    for coord in coords:
        if all(math.isnan(val) for val in coord):
            continue
        x, y, z = coord
        point = App.Vector(x, y, z)
        Draft.makePoint(point, color=(0.0, 1.0, 0.0))

# Predict a circle from the circumfance
def create_circle_from_points(coords):

    # Filter out coordinates with NaN values
    valid_coords = [coord for coord in coords if not any(math.isnan(val) for val in coord)]

    if not valid_coords:
        print("No valid coordinates to create circle.")
        return

    # Predict circle parameters
    points = np.array(valid_coords)
    initial_guess = [np.mean(points[:, 0]), np.mean(points[:, 1]), np.mean(points[:, :2].ptp(axis=0))]

    def circle_residuals(params, points):
        cx, cy, r = params
        distances = np.sqrt((points[:, 0] - cx) ** 2 + (points[:, 1] - cy) ** 2)
        return np.sum((distances - r) ** 2)
   
    result = minimize(circle_residuals, initial_guess, args=(points,), method='Nelder-Mead')

    # Define circle parameters 
    cx, cy, r = result.x
    cz = np.mean(points[:, 2])

    # Create circle
    circle_wire = Draft.makeCircle(radius=r, placement=App.Placement(App.Vector(cx, cy, cz), App.Rotation(App.Vector(0, 0, 1), 0)), face=False)

    return cx, cy, cz, r

def create_cylinder_from_circles(circle1_params, circle2_params):
    global No_Go_Zone_Radius

    # Extract parameters of circles
    cx1, cy1, cz1, r1 = circle1_params
    cx2, cy2, cz2, r2 = circle2_params

    # Calculate cylinder height
    height = abs(cz2 - cz1)
    # Calculate the center of the cylinder
    cx = (cx1 + cx2) / 2
    cy = (cy1 + cy2) / 2

    # Create outer cylinder
    outer_cylinder = Part.makeCylinder(r1, height, App.Vector(cx, cy, min(cz1, cz2)))
    # Create inner cylinder
    inner_radius = 0.99 * r1  # Adjust the scale factor as needed for the inner radius
    inner_cylinder = Part.makeCylinder(inner_radius, height, App.Vector(cx, cy, min(cz1, cz2)))

    # Make the outer cylinder hollow by subtracting the inner cylinder
    hollow_cylinder = outer_cylinder.cut(inner_cylinder)

    # Create the hollow cylinder (representing BTST)
    Part.show(hollow_cylinder)

    # Create the No Go Zone with transparency
    No_Go_Zone_Radius = 1.148
    No_Go_Zone = Part.makeCylinder(No_Go_Zone_Radius, height, App.Vector(0,0,min(cz1, cz2)))
    
    # Show the No Go Zone
    Part.show(No_Go_Zone)
    
    # Apply transparency and colour to the No Go Zone (might need to edit to define a name)
    Gui.ActiveDocument.getObject("Shape001").Transparency = 80
    Gui.ActiveDocument.getObject("Shape001").ShapeColor = (1.0, 0.0, 0.0)

# Extract Circumference data and convert to real numbers
Z_plus_end = df.iloc[1:, 1:4].apply(pd.to_numeric, errors='coerce').values.tolist()
Z_minus_end = df.iloc[1:, 6:9].apply(pd.to_numeric, errors='coerce').values.tolist()

# Plot the data points for the Circumfrance
create_points(Z_plus_end)
create_points(Z_minus_end)

# Create circles from points
circle1_params = create_circle_from_points(Z_plus_end)
circle2_params = create_circle_from_points(Z_minus_end)

# Create cylinder from circles
if circle1_params and circle2_params:
    create_cylinder_from_circles(circle1_params, circle2_params)

coord_index = []
# Mark the locations of measurements along the I-beams
def mark_points(coords):
    global radius, separation, coord_index

    radius = [] ; separation = [] 

    # Filter out coordinates with NaN values
    valid_coords = [coord for coord in coords if not any(math.isnan(val) for val in coord)]
    if not valid_coords:
        return

    #Calculate radius and lateral displacement
    for i in range(0, 7):
        coordinates = [coords[i][0],coords[i][1],coords[i][2]]
        coord_index.append(coordinates)

        calc_radius = math.sqrt(((coords[i][0]) ** 2) + ((coords[i][1]) ** 2)) # Measured radius
        calc_radius += 0.01442 # Adjusted radius

        # Store radius
        if calc_radius > 0:
            radius.append(calc_radius)

        # Calculate and store separation distance
        if len(coord_index) > 7:
            calc_separation = math.sqrt((coord_index[-8][0] - coords[i][0])**2 + (coord_index[-8][1] - coords[i][1])**2 + (coord_index[-8][2] - coords[i][2])**2)
            separation.append(calc_separation)

    #Create the point's along the I-beam
    for coord, rad in zip(coords, radius):
        if all(math.isnan(val) for val in coord):
            continue
        x, y, z = coord

        # Calculated angle based on the point's position
        angle = math.atan2(y, x)

        # Calculated new position based on polar coordinates
        new_x = rad * math.cos(angle)
        new_y = rad * math.sin(angle)
        
        point = App.Vector(new_x, new_y, z)
 
        #Measured point
        #Draft.makePoint(x,y,z,color=(1.0,0.0,0.0))

        #Adjusted point
        Draft.makePoint(point, color=(0.0, 0.0, 1.0))

index = 0 ; index_sign = "+" ; matrix = []

#Calculate necessary adjustments vertically and laterally
def check_points(coords):
    global radius, separation, index, index_sign
    
    radial_index = [] ; rad_adjust_index = []
    separation_index = [] ; sep_adjust_index = []

    nominal_separation = 0.184 + 0.002366
    index += 1 

    if index > 38:
        index_sign = "-"
        index = index - 38 ; index_val=index_sign+str(index)
        radial_index.append(index_val) ; rad_adjust_index.append(index_val)
        separation_index.append(index_val) ; sep_adjust_index.append(index_val)
    else:
        index_val = index_sign+str(index)
        radial_index.append(index_val) ; rad_adjust_index.append(index_val)
        separation_index.append(index_val) ; sep_adjust_index.append(index_val)

    for i in range(0,7):
        #Calculate necessary adjustment vertically
        if len(radius) > 0 and radius[i] < No_Go_Zone_Radius:
            radial_difference = No_Go_Zone_Radius - radius[i] ; radial_difference *= 1000 ; radial_difference = round(radial_difference,2)
            radial_difference = str(radial_difference) + "mm"
            radial_index.append(radial_difference) ; rad_adjust_index.append(str(i+1))
        else:
            radial_difference = "0mm" ; radial_index.append(radial_difference) ; rad_adjust_index.append("-")

        #Calculate necessary adjustment laterally
        if len(separation) > 0 and separation[i] != nominal_separation: #to be editing to within tolerance [DEFINE TOLERANCE ON WEDNESDAY]
            separation_difference = separation[i] - nominal_separation ; separation_difference *= 1000 ; separation_difference = round(separation_difference,3)
            adjust_separation =  separation_difference * -1 ; adjust_separation = str(adjust_separation) + "mm"
            separation_index.append(adjust_separation) ; sep_adjust_index.append(str(i+1))
        else:
            separation_difference = "0mm" ; separation_index.append(separation_difference) ; sep_adjust_index.append("-")  

    # Also measures the seperation between 19 and 20 for +/- which isn't necessary ; to be edited!

    matrix.append(rad_adjust_index) ; matrix.append(radial_index)
    matrix.append(sep_adjust_index) ; matrix.append(separation_index)
            
 
# Extract I-beam data and convert to real numbers
I_plus_1 = df.iloc[1:8, 11:14].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_2 = df.iloc[1:8, 16:19].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_3 = df.iloc[1:8, 21:24].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_4 = df.iloc[1:8, 26:29].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_5 = df.iloc[1:8, 31:34].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_6 = df.iloc[1:8, 36:39].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_7 = df.iloc[1:8, 41:44].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_8 = df.iloc[1:8, 46:49].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_9 = df.iloc[1:8, 51:54].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_10 = df.iloc[1:8, 56:59].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_11 = df.iloc[1:8, 61:64].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_12 = df.iloc[1:8, 66:69].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_13 = df.iloc[1:8, 71:74].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_14 = df.iloc[1:8, 76:79].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_15 = df.iloc[1:8, 81:84].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_16 = df.iloc[1:8, 86:89].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_17 = df.iloc[1:8, 91:94].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_18 = df.iloc[1:8, 96:99].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_19 = df.iloc[1:8, 101:104].apply(pd.to_numeric, errors='coerce').values.tolist()

I_plus_20 = df.iloc[11:18, 11:14].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_21 = df.iloc[11:18, 16:19].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_22 = df.iloc[11:18, 21:24].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_23 = df.iloc[11:18, 26:29].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_24 = df.iloc[11:18, 31:34].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_25 = df.iloc[11:18, 36:39].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_26 = df.iloc[11:18, 41:44].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_27 = df.iloc[11:18, 46:49].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_28 = df.iloc[11:18, 51:54].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_29 = df.iloc[11:18, 56:59].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_30 = df.iloc[11:18, 61:64].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_31 = df.iloc[11:18, 66:69].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_32 = df.iloc[11:18, 71:74].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_33 = df.iloc[11:18, 76:79].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_34 = df.iloc[11:18, 81:84].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_35 = df.iloc[11:18, 86:89].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_36 = df.iloc[11:18, 91:94].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_37 = df.iloc[11:18, 96:99].apply(pd.to_numeric, errors='coerce').values.tolist()
I_plus_38 = df.iloc[11:18, 101:104].apply(pd.to_numeric, errors='coerce').values.tolist()

I_minus_1 = df.iloc[21:28, 11:14].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_2 = df.iloc[21:28, 16:19].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_3 = df.iloc[21:28, 21:24].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_4 = df.iloc[21:28, 26:29].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_5 = df.iloc[21:28, 31:34].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_6 = df.iloc[21:28, 36:39].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_7 = df.iloc[21:28, 41:44].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_8 = df.iloc[21:28, 46:49].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_9 = df.iloc[21:28, 51:54].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_10 = df.iloc[21:28, 56:59].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_11 = df.iloc[21:28, 61:64].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_12 = df.iloc[21:28, 66:69].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_13 = df.iloc[21:28, 71:74].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_14 = df.iloc[21:28, 76:79].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_15 = df.iloc[21:28, 81:84].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_16 = df.iloc[21:28, 86:89].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_17 = df.iloc[21:28, 91:94].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_18 = df.iloc[21:28, 96:99].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_19 = df.iloc[21:28, 101:104].apply(pd.to_numeric, errors='coerce').values.tolist()

I_minus_20 = df.iloc[31:38, 11:14].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_21 = df.iloc[31:38, 16:19].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_22 = df.iloc[31:38, 21:24].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_23 = df.iloc[31:38, 26:29].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_24 = df.iloc[31:38, 31:34].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_25 = df.iloc[31:38, 36:39].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_26 = df.iloc[31:38, 41:44].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_27 = df.iloc[31:38, 46:49].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_28 = df.iloc[31:38, 51:54].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_29 = df.iloc[31:38, 56:59].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_30 = df.iloc[31:38, 61:64].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_31 = df.iloc[31:38, 66:69].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_32 = df.iloc[31:38, 71:74].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_33 = df.iloc[31:38, 76:79].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_34 = df.iloc[31:38, 81:84].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_35 = df.iloc[31:38, 86:89].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_36 = df.iloc[31:38, 91:94].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_37 = df.iloc[31:38, 96:99].apply(pd.to_numeric, errors='coerce').values.tolist()
I_minus_38 = df.iloc[31:38, 101:104].apply(pd.to_numeric, errors='coerce').values.tolist()

# Plot the data points for the I-beams in z+
mark_points(I_plus_1) ; check_points(I_plus_1)
mark_points(I_plus_2) ; check_points(I_plus_2)
mark_points(I_plus_3) ; check_points(I_plus_3)
mark_points(I_plus_4) ; check_points(I_plus_4)
mark_points(I_plus_5) ; check_points(I_plus_5)
mark_points(I_plus_6) ; check_points(I_plus_6)
mark_points(I_plus_7) ; check_points(I_plus_7)
mark_points(I_plus_8) ; check_points(I_plus_8)
mark_points(I_plus_9) ; check_points(I_plus_9)
mark_points(I_plus_10) ; check_points(I_plus_10)
mark_points(I_plus_11) ; check_points(I_plus_11)
mark_points(I_plus_12) ; check_points(I_plus_12)
mark_points(I_plus_13) ; check_points(I_plus_13)
mark_points(I_plus_14) ; check_points(I_plus_14)
mark_points(I_plus_15) ; check_points(I_plus_15)
mark_points(I_plus_16) ; check_points(I_plus_16)
mark_points(I_plus_17) ; check_points(I_plus_17)
mark_points(I_plus_18) ; check_points(I_plus_18)
mark_points(I_plus_19) ; check_points(I_plus_19)
mark_points(I_plus_20) ; check_points(I_plus_20)
mark_points(I_plus_21) ; check_points(I_plus_21)
mark_points(I_plus_22) ; check_points(I_plus_22)
mark_points(I_plus_23) ; check_points(I_plus_23)
mark_points(I_plus_24) ; check_points(I_plus_24)
mark_points(I_plus_25) ; check_points(I_plus_25)
mark_points(I_plus_26) ; check_points(I_plus_26)
mark_points(I_plus_27) ; check_points(I_plus_27)
mark_points(I_plus_28) ; check_points(I_plus_28)
mark_points(I_plus_29) ; check_points(I_plus_29)
mark_points(I_plus_30) ; check_points(I_plus_30)
mark_points(I_plus_31) ; check_points(I_plus_31)
mark_points(I_plus_32) ; check_points(I_plus_32)
mark_points(I_plus_33) ; check_points(I_plus_33)
mark_points(I_plus_34) ; check_points(I_plus_34)
mark_points(I_plus_35) ; check_points(I_plus_35)
mark_points(I_plus_36) ; check_points(I_plus_36)
mark_points(I_plus_37) ; check_points(I_plus_37)
mark_points(I_plus_38) ; check_points(I_plus_38)

# Plot the data points for the I-beams in z-
mark_points(I_minus_1) ; check_points(I_minus_1)
mark_points(I_minus_2) ; check_points(I_minus_2)
mark_points(I_minus_3) ; check_points(I_minus_3)
mark_points(I_minus_4) ; check_points(I_minus_4)
mark_points(I_minus_5) ; check_points(I_minus_5)
mark_points(I_minus_6) ; check_points(I_minus_6)
mark_points(I_minus_7) ; check_points(I_minus_7)
mark_points(I_minus_8) ; check_points(I_minus_8)
mark_points(I_minus_9) ; check_points(I_minus_9)
mark_points(I_minus_10) ; check_points(I_minus_10)
mark_points(I_minus_11) ; check_points(I_minus_11)
mark_points(I_minus_12) ; check_points(I_minus_12)
mark_points(I_minus_13) ; check_points(I_minus_13)
mark_points(I_minus_14) ; check_points(I_minus_14)
mark_points(I_minus_15) ; check_points(I_minus_15)
mark_points(I_minus_16) ; check_points(I_minus_16)
mark_points(I_minus_17) ; check_points(I_minus_17)
mark_points(I_minus_18) ; check_points(I_minus_18)
mark_points(I_minus_19) ; check_points(I_minus_19)
mark_points(I_minus_20) ; check_points(I_minus_20)
mark_points(I_minus_21) ; check_points(I_minus_21)
mark_points(I_minus_22) ; check_points(I_minus_22)
mark_points(I_minus_23) ; check_points(I_minus_23)
mark_points(I_minus_24) ; check_points(I_minus_24)
mark_points(I_minus_25) ; check_points(I_minus_25)
mark_points(I_minus_26) ; check_points(I_minus_26)
mark_points(I_minus_27) ; check_points(I_minus_27)
mark_points(I_minus_28) ; check_points(I_minus_28)
mark_points(I_minus_29) ; check_points(I_minus_29)
mark_points(I_minus_30) ; check_points(I_minus_30)
mark_points(I_minus_31) ; check_points(I_minus_31)
mark_points(I_minus_32) ; check_points(I_minus_32)
mark_points(I_minus_33) ; check_points(I_minus_33)
mark_points(I_minus_34) ; check_points(I_minus_34)
mark_points(I_minus_35) ; check_points(I_minus_35)
mark_points(I_minus_36) ; check_points(I_minus_36)
mark_points(I_minus_37) ; check_points(I_minus_37)
mark_points(I_minus_38) ; check_points(I_minus_38)

#Adjustment pop-up (informing us how much we should adjust the feet by)
def adjustment():
    global matrix

    window = Tk()
    window.title("Adjustment Required")
    window.geometry("1650x200+120+700")
    
    canvas = Canvas(window, bg="#000000")
    canvas.pack(side="left", fill="both", expand=True)

    #Create a scroll bar
    scrollbar = Scrollbar(window, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    frame = Frame(canvas, bg="#000000")
    canvas.create_window((0, 0), window=frame, anchor="nw")

    Label(frame,text = "",font=("Lato",30,"bold"),fg="#fff", bg="#000000").grid(row=0,column=0)
    Label(frame,text = "Adjustments Required",font=("Lato",20,"bold"), fg="#fff", bg="#000000").place(x=650,y=8)

    index = 0 ; sep_index = 2

    for i in range(0, 77):
        for j in range(1, 8):
            if matrix[index][j] != "-":
                if int(matrix[index][0]) > 0:
                    display_text = f"Foot {matrix[index][j]} by >{matrix[index+1][j]} vertically"
                else:
                    display_text = f"Foot -{matrix[index][j]} by >{matrix[index+1][j]} vertically"
                Label(frame, text=f"     I-Beam {matrix[index][0]} ",font=("Perpetua",10,"bold"), fg="#fff", bg="#000000").grid(row=(2*i+1), column=0)
                Label(frame, text="|", fg="#fff", bg="#000000").grid(row=(2*i+1), column=(2*j))
                Label(frame, text=display_text, font=("Lato",8,"bold"), fg="#fff", bg="#000000").grid(row=(2*i+1), column=(2*j + 1))

            if matrix[sep_index][j] != "-":
                if int(matrix[sep_index][0]) > 0:
                    display_text = f"Foot {matrix[sep_index][j]} by {matrix[sep_index+1][j]} laterally"
                else:
                    display_text = f"Foot -{matrix[sep_index][j]} by {matrix[sep_index+1][j]} laterally"
                Label(frame, text=f"     I-Beam {matrix[sep_index-4][0]} ~ {matrix[sep_index][0]} ",font=("Perpetua",10,"bold"), fg="#fff", bg="#000000").grid(row=2*i, column=0)
                Label(frame, text="|", fg="#fff", bg="#000000").grid(row=2*i, column=(2*j))
                Label(frame, text=display_text, font=("Lato",8,"bold"), fg="#fff", bg="#000000").grid(row=2*i, column=(2*j + 1))
        if index < 300:
            index += 4
        else:
            break  
        if sep_index < 302:
            sep_index += 4
        else:
            break 
    
    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    window.mainloop()  

App.ActiveDocument.recompute()

# Open the adjustment pop-up
adjustment()
