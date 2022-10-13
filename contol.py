#Import tkinter library
from tkinter import *
import PIL
from PIL import Image,ImageTk
import pytesseract
import tkinter as tk
from PIL import Image, ImageTk
import numpy as np
import math
import cv2
import numpy as np
import math
from tkinter import ttk
import cv2
import numpy as np
from imutils import perspective
import sys
import cv2 as cv
import time
from scipy.spatial.distance import euclidean
from openpyxl import Workbook,load_workbook
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import PIL
from PIL import Image,ImageTk
import keyboard
import webbrowser

wb = load_workbook("MENSA.xlsx")
ws = wb.active


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("PDF files","*.pdf"),("All Files", "*.*")))
    label_file["text"] = filename
    """
    root.update()
    """
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clear_text():
   text.delete(1.0, END)
def clear_data():
    tv1.delete(*tv1.get_children())
    return None
def empty(a):
    pass

t=0

cv2.namedWindow("Parameters")
cv2.resizeWindow("Parameters",640,240)
cv2.createTrackbar("Threshold1","Parameters",23,255,empty)
cv2.createTrackbar("Threshold2","Parameters",20,255,empty)
cv2.createTrackbar("Area","Parameters",5000,30000,empty)

#Create an instance of tkinter frame or window
win= Tk()
#Set the geometry of tkinter frame
win.geometry("1500x800")
win.title("Loadcell Control")
#Define a new function to open the window



def getContours(img, imgContour):
    contours, hierarchy = cv2.findContours(img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    if len(contours) != 0:
        ref_object = contours[0]
        box = cv2.minAreaRect(ref_object)
        box = cv2.boxPoints(box)
        box = np.array(box, dtype="int")
        box = perspective.order_points(box)
        (tl, tr, br, bl) = box
        dist_in_pixel = euclidean(tl, tr)
        dist_in_cm = 2
        if t == 0:
            pixel_per_cm = 1
        else:
            pixel_per_cm = t
        # pixel_per_cm = 0.0045662100456621
        a = 0
        for cnt in contours:
            area = cv2.contourArea(cnt)
            areaMin = cv2.getTrackbarPos("Area", "Parameters")
            if area > areaMin:
                box = cv2.minAreaRect(cnt)
                box = cv2.boxPoints(box)
                box = np.array(box, dtype="int")
                box = perspective.order_points(box)
                (tl, tr, br, bl) = box
                cv2.drawContours(imgContour, cnt, -1, (255, 0, 255), 7)

                mid_pt_horizontal = (tl[0] + int(abs(tr[0] - tl[0]) / 2), tl[1] + int(abs(tr[1] - tl[1]) / 2))
                mid_pt_verticle = (tr[0] + int(abs(tr[0] - br[0]) / 2), tr[1] + int(abs(tr[1] - br[1]) / 2))
                wid = euclidean(tl, tr) * pixel_per_cm
                ht = euclidean(tr, br) * pixel_per_cm
                a = a + 1
                # print(wid*pixel_per_cm)
                # print(ht*pixel_per_cm)

                peri = cv2.arcLength(cnt, True)
                approx = cv2.approxPolyDP(cnt, 0.02 * peri, True)  # kontur sayısı toplam
                print(len(approx))  # EXCEL TABLO ISLEMLERI BURADA YAPILACAK

                p = [1, 2, 3, 4, 5]
                x = []


                x, y, w, h = cv2.boundingRect(approx)
                if keyboard.is_pressed('p'):
                    calib(w, h)
                    break
                if cv2.waitKey(1) & 0xFF == ord('w'):
                    print("riza")
                    calib(w, h)

                cv2.rectangle(imgContour, (x, y), (x + w, y + h), (0, 255, 0), 5)

                cv2.putText(imgContour, "Points: " + str(len(approx)), (x + w + 20, y + 20), cv2.FONT_HERSHEY_COMPLEX, .7,
                            (0, 255, 0), 2)
                cv2.putText(imgContour, "Area: " + str(int(area)), (x + w + 20, y + 45), cv2.FONT_HERSHEY_COMPLEX, 0.7,
                            (0, 255, 0), 2)

                cv2.putText(imgContour, "{:.1f}cm".format(wid),
                            (int(mid_pt_horizontal[0] - 15), int(mid_pt_horizontal[1] - 10)), cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                            (255, 255, 0), 2)
                cv2.putText(imgContour, "{:.1f}cm".format(ht), (int(mid_pt_verticle[0] + 10), int(mid_pt_verticle[1])),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)




def stackImages(scale,imgArray):
    rows = len(imgArray)
    cols = len(imgArray[0])
    rowsAvailable = isinstance(imgArray[0], list)
    width = imgArray[0][0].shape[1]
    height = imgArray[0][0].shape[0]
    if rowsAvailable:
        for x in range ( 0, rows):
            for y in range(0, cols):
                if imgArray[x][y].shape[:2] == imgArray[0][0].shape [:2]:
                    imgArray[x][y] = cv2.resize(imgArray[x][y], (0, 0), None, scale, scale)
                else:
                    imgArray[x][y] = cv2.resize(imgArray[x][y], (imgArray[0][0].shape[1], imgArray[0][0].shape[0]), None, scale, scale)
                if len(imgArray[x][y].shape) == 2: imgArray[x][y]= cv2.cvtColor( imgArray[x][y], cv2.COLOR_GRAY2BGR)
        imageBlank = np.zeros((height, width, 3), np.uint8)
        hor = [imageBlank]*rows
        hor_con = [imageBlank]*rows
        for x in range(0, rows):
            hor[x] = np.hstack(imgArray[x])
        ver = np.vstack(hor)
    else:
        for x in range(0, rows):
            if imgArray[x].shape[:2] == imgArray[0].shape[:2]:
                imgArray[x] = cv2.resize(imgArray[x], (0, 0), None, scale, scale)
            else:
                imgArray[x] = cv2.resize(imgArray[x], (imgArray[0].shape[1], imgArray[0].shape[0]), None,scale, scale)
            if len(imgArray[x].shape) == 2: imgArray[x] = cv2.cvtColor(imgArray[x], cv2.COLOR_GRAY2BGR)
        hor= np.hstack(imgArray)
        ver = hor
    return ver



class MainWindow():
   def __init__(self, window, cap):
      self.window = window
      self.cap = cap
      self.width = self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
      self.height = self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
      self.interval = 20  # Interval in ms to get the latest frame
      # Create canvas for image
      self.canvas = tk.Canvas(self.window, width=1500, height=900)
      self.canvas.grid(row=0, column=0)
      # Update image on canvas
      self.update_image()


   def update_image(self):

      self.image = self.cap.read()[1]

      self.polepoz = np.split(self.image, 2, axis=1)
      self.imgContour = self.polepoz[0].copy()
      self.imgBlur = cv2.GaussianBlur(self.image, (7, 7), 1)
      self.imgGray = cv2.cvtColor(self.imgBlur, cv2.COLOR_BGR2GRAY)
      # imgContour = imgGray.copy()
      self.threshold1 = cv2.getTrackbarPos("Threshold1", "Parameters")
      self.threshold2 = cv2.getTrackbarPos("Threshold2", "Parameters")
      self.imgCanny = cv2.Canny(self.imgGray, self.threshold1, self.threshold2)
      self.kernel = np.ones((5, 5))
      self.imgDil = cv2.dilate(self.imgCanny, self.kernel, iterations=1)
      getContours(self.imgDil, self.imgContour)
      self.imgStack = stackImages(0.8, ([self.image, self.imgCanny],
                                   [self.imgDil, self.imgContour]))

      self.imgContour = Image.fromarray(self.imgContour)  # to PIL format
      self.imgContour = ImageTk.PhotoImage(self.imgContour)  # to ImageTk format
      # Update image
      self.canvas.create_image(0, 0, anchor=tk.NW, image=self.imgContour)
      # Repeat every 'interval' ms
      self.window.after(self.interval, self.update_image)

def web():
    name= filedialog.askopenfilename()
    print(name)
    print(type(name))
    webbrowser.open_new(name)
    #webbrowser.open_new_tab(name)
    #webbrowser.get().open('name')
    #return None
def excelac():
    import os
    name=filedialog.askopenfilename()
    os.system(name)

def open_win():

   #_, frame = cap.read()


   new= Toplevel(win)
   new.geometry("1500x900")
   new.title("Running Program")
   MainWindow(new, cv2.VideoCapture(0))







"""
---------------------------------------
               Main Window
---------------------------------------
"""





#Create a label
ll=Label(win, text= "L-01-01Parcasi :", font= ('Helvetica 11 bold'))
ll.place(x=40, y=50)
#ll.place(relx=0.0, rely=1.0, anchor='sw')
#Create a button to open a New Window
but=ttk.Button(win, text="Calistir", command=open_win)
but.place(x=180,y=50)

ll=Label(win, text="L-01-02 Parcasi :", font= ('Helvetica 11 bold'))
ll.place(x=40, y=80)
#ll.place(relx=0.0, rely=1.0, anchor='sw')
#Create a button to open a New Window
but=ttk.Button(win, text="Calistir", command=open_win)
but.place(x=180,y=80)
"""
button3=tk.Button(win,text="open",command=lambda: openNewWindow())
button3.place(rely=0.65, relx=0.30)   
"""
# Frame for TreeView
frame1 = LabelFrame(win, text="Excel Data")
frame1.place(height=250, width=500)

# Frame for open file dialog
file_frame = LabelFrame(win, text="Open File")
file_frame.place(height=100, width=400, rely=0.65, relx=0)

# Buttons
button1 = Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.65, relx=0.50)

button2 = Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.30)

ll=Label(win, text= "PDF BAK :", font= ('Helvetica 11 bold'))
ll.place(x=640, y=170)
button3=Button(win,text="open PDF",command=lambda: web())
button3.place(x=780,y=170) 

ll=Label(win, text= "Excel BAK :", font= ('Helvetica 11 bold'))
ll.place(x=640, y=220)
button3=Button(win,text="open Excel",command=lambda: excelac())
button3.place(x=780,y=220)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)


## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).
ll=Label(win, text= "Camera", font= ('Helvetica 11 bold'))
ll.place(x=640, y=110)
#ll.place(relx=0.0, rely=1.0, anchor='sw')
#Create a button to open a New Window
but=ttk.Button(win, text="Calistir", command=open_win)
but.place(x=780,y=110)


ll=Label(win, text= "L-01-04 Parcasi :", font= ('Helvetica 11 bold'))
ll.place(x=640, y=140)
#ll.place(relx=0.0, rely=1.0, anchor='sw')
#Create a button to open a New Window
but=ttk.Button(win, text="Calistir:", command=open_win)
but.place(x=780,y=140)


"""
Mainloop function
"""
win.mainloop()