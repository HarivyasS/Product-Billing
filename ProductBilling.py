"""
Created on Wed Dec  2 19:20:14 2020

@author: Harivyas
"""
import tkinter as tk
from tkinter import Canvas
from scipy.spatial.distance import euclidean
from imutils import perspective
from imutils import contours
import numpy as np
import imutils
import cv2
import os
import xlwt 
from xlwt import Workbook 
from PIL import ImageTk

i=0
i=i+1

a=0
total=0
wb = Workbook() 

style = xlwt.easyxf('font: bold 1') 
color = xlwt.easyxf('font: color red') 
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1') 

sheet1.write(0, 0, 'Sr. No',style) 
sheet1.write(0, 1, 'Product ID',style) 
sheet1.write(0, 2, 'Name of Product',style) 
sheet1.write(0, 3, 'Quantity ',style) 
sheet1.write(0, 4, 'Cost of Product',style)
sheet1.write(0, 5, 'Final Cost of Product',style) 


def camera():
    
    message = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
    message.place(x=500, y=200)
   
    cam = cv2.VideoCapture(0)
    cv2.namedWindow("Press Enter to Capture")
   
    img_counter = 0
   
    while True:
        ret, frame = cam.read()
        cv2.imshow("Press Enter to Capture", frame)
        if not ret:
            break
        k = cv2.waitKey(1)
        #print(k)
        if k%256 == 27:
            # ESC pressed
            print("Escape hit, closing...")
            break
        elif k%256 == 13:
            # SPACE pressed
            img_name = "C:/Users/asd/Desktop/sample/newimage.jpg".format(img_counter)
            cv2.imwrite(img_name, frame)
            print("{} written!".format(img_name))
            img_counter += 1
            
    cam.release()
   
    cv2.destroyAllWindows()
    
def show_images(images):
    for c, img in enumerate(images):
        cv2.imshow("image_" + str(i), img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    img_path = "C:/Users/asd/Desktop/sample/newimage.jpg"
    
    # Read image and preprocess
    image = cv2.imread(img_path)
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (9, 9), 0)
    
    edged = cv2.Canny(blur, 50, 100)
    edged = cv2.dilate(edged, None, iterations=1)
    edged = cv2.erode(edged, None, iterations=1)
    
    #show_images([blur, edged])
    
    # Find contours
    cnts = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    
    # Sort contours from left to right as leftmost contour is reference object
    (cnts, _) = contours.sort_contours(cnts)
    
    # Remove contours which are not large enough
    cnts = [x for x in cnts if cv2.contourArea(x) > 100]
    
    cv2.drawContours(image, cnts, -1, (0,255,0), 3)
    
    show_images([image, edged])
    print(len(cnts))
    
    # Reference object dimensions
    # Here for reference I have used a 2cm x 2cm square
    ref_object = cnts[0]
    box = cv2.minAreaRect(ref_object)
    box = cv2.boxPoints(box)
    box = np.array(box, dtype="int")
    box = perspective.order_points(box)
    (tl, tr, br, bl) = box
    dist_in_pixel = euclidean(tl, tr)
    dist_in_cm = 2
    pixel_per_cm = dist_in_pixel/dist_in_cm
    
    for cnt in cnts:
        box = cv2.minAreaRect(cnt)
        box = cv2.boxPoints(box)
        box = np.array(box, dtype="int")
        box = perspective.order_points(box)
        (tl, tr, br, bl) = box
        cv2.drawContours(image, [box.astype("int")], -1, (0, 0, 255), 2)
        mid_pt_horizontal = (tl[0] + int(abs(tr[0] - tl[0])/2), tl[1] + int(abs(tr[1] - tl[1])/2))
        mid_pt_verticle = (tr[0] + int(abs(tr[0] - br[0])/2), tr[1] + int(abs(tr[1] - br[1])/2))
        wid = euclidean(tl, tr)/pixel_per_cm
        ht = euclidean(tr, br)/pixel_per_cm  
        a="%.1f" % ht
        print(a)
        cv2.putText(image, "{:.1f}cm".format(wid), (int(mid_pt_horizontal[0] - 15), int(mid_pt_horizontal[1] - 10)),cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)
        cv2.putText(image, "{:.1f}cm".format(ht), (int(mid_pt_verticle[0] + 10), int(mid_pt_verticle[1])),cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)
        break;
    
    if(float(a)==0.9 or float(a)==1.0 or float(a)==0.5 or float(a)==0.4 or float(a)==1.4 or float(a)==0.6 ):
        message = tk.Label(root, text="Detected" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
        message.place(x=500, y=200)
        
        #mess = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
        mess.place(x=500, y=250)

        
    else:
        message = tk.Label(root, text="NOT Detected" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
        message.place(x=500, y=200) 
        
        #mess = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
        mess.place(x=500, y=250)
     
    

def bill():
    global i
    #global a
    global total
    
        
    def show_images(images):
        for c, img in enumerate(images):
          cv2.imshow("image_" + str(i), img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
    img_path = "C:/Users/asd/Desktop/sample/new.jpg"
    
    # Read image and preprocess
    image = cv2.imread(img_path)
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (9, 9), 0)
    
    edged = cv2.Canny(blur, 50, 100)
    edged = cv2.dilate(edged, None, iterations=1)
    edged = cv2.erode(edged, None, iterations=1)
    
    #show_images([blur, edged])
    
    # Find contours
    cnts = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    
    # Sort contours from left to right as leftmost contour is reference object
    (cnts, _) = contours.sort_contours(cnts)
    
    # Remove contours which are not large enough
    cnts = [x for x in cnts if cv2.contourArea(x) > 100]
    
    #cv2.drawContours(image, cnts, -1, (0,255,0), 3)
    
    #show_images([image, edged])
    #print(len(cnts))
    
    # Reference object dimensions
    # Here for reference I have used a 2cm x 2cm square
    ref_object = cnts[0]
    box = cv2.minAreaRect(ref_object)
    box = cv2.boxPoints(box)
    box = np.array(box, dtype="int")
    box = perspective.order_points(box)
    (tl, tr, br, bl) = box
    dist_in_pixel = euclidean(tl, tr)
    dist_in_cm = 2
    pixel_per_cm = dist_in_pixel/dist_in_cm
    
    for cnt in cnts:
        box = cv2.minAreaRect(cnt)
        box = cv2.boxPoints(box)
        box = np.array(box, dtype="int")
        box = perspective.order_points(box)
        (tl, tr, br, bl) = box
        cv2.drawContours(image, [box.astype("int")], -1, (0, 0, 255), 2)
        mid_pt_horizontal = (tl[0] + int(abs(tr[0] - tl[0])/2), tl[1] + int(abs(tr[1] - tl[1])/2))
        mid_pt_verticle = (tr[0] + int(abs(tr[0] - br[0])/2), tr[1] + int(abs(tr[1] - br[1])/2))
        wid = euclidean(tl, tr)/pixel_per_cm
        ht = euclidean(tr, br)/pixel_per_cm  
        a="%.1f" % ht
        #print(a)
        cv2.putText(image, "{:.1f}cm".format(wid), (int(mid_pt_horizontal[0] - 15), int(mid_pt_horizontal[1] - 10)),cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)
        cv2.putText(image, "{:.1f}cm".format(ht), (int(mid_pt_verticle[0] + 10), int(mid_pt_verticle[1])),cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)
        break;

    
    
    
    quantity=(txt.get())
    
    
    
    if(float(a)==0.9):
        
        sheet1.write(i, 0, i)
        cost=int(54)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '111')
        sheet1.write(i, 2, 'Gillette')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    elif(float(a)==1.0):
        
        sheet1.write(i, 0, i)
        cost=int(168)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '222')
        sheet1.write(i, 2, 'NOSIKIND-nasal solution')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    elif(float(a)==0.5):
        
        sheet1.write(i, 0, i)
        cost=int(10)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '333')
        sheet1.write(i, 2, 'colgate small')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    elif(float(a)==0.4):
        sheet1.write(i, 0, i)
        cost=int(55)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '444')
        sheet1.write(i, 2, 'colgate large')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    elif(float(a)==1.4):
        sheet1.write(i, 0, i)
        cost=int(10)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '555')
        sheet1.write(i, 2, 'SOAP-Mysore Sandal')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    elif(float(a)==0.6):
        sheet1.write(i, 0, i)
        cost=int(10)
        amount=cost*int(quantity)
        sheet1.write(i, 1, '666')
        sheet1.write(i, 2, 'Vovenac gel')
        sheet1.write(i, 3, quantity)
        sheet1.write(i, 4, cost)
        sheet1.write(i, 5, amount)
        i=i+1
        total=total+amount
    
   
    message = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
    message.place(x=500, y=200)
    
    txt.delete('0')
    
    mess = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
    mess.place(x=500, y=250)

    
def display():
    
    sheet1.write(i, 5, total,color)
    sheet1.write(i, 6, 'FINAL COST',color)
    wb.save('C:/Users/asd/Desktop/sample/excel.xls')  
    os.system('start "excel" "C:/Users/asd/Desktop/sample/excel.xls"')
    

    

root = tk.Tk()
canvas= Canvas(width = 1000, height = 1000, bg = 'white')
canvas.pack(expand = 'YES', fill = 'both')

#image = ImageTk.PhotoImage(file = "C:/Users/asd/Desktop/sample/clg.png")
#canvas.create_image(500, 40, image = image, anchor = 'nw')

lbl = tk.Label(root, text="TEAM NAME : TEAM #INCLUDE ",width=50  ,height=0  ,fg="blue"  ,bg="white" ,font=('Copperplate Gothic Bold', 15, ' bold ') )
lbl.place(x=100, y=50)



capture = tk.Button(root,
                   text="Capture",fg="black"  ,bg="#00CED1",
                    width=41 ,height=2,activebackground = "Red" ,font=('Comic Sans MS', 15, ' bold '),
                   command=camera)
capture.place(x=230,y=100)

txt=tk.Entry(root,width=20  ,bg="silver" ,fg="black",font=('times', 18, ' bold '),borderwidth=2,relief="solid")        
txt.place(x=500, y=250)

lbl3 = tk.Label(root, text="Notification : ",width=19  ,height=0  ,fg="black"  ,bg="#00CED1"  ,font=('Comic Sans MS', 15, ' bold'))
lbl3.place(x=230, y=200)

message = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=1,relief="solid")
message.place(x=500, y=200)

lbl = tk.Label(root, text="Enter Quantity :",width=19  ,height=0  ,fg="black"  ,bg="#00CED1" ,font=('Comic Sans MS', 15, ' bold ') )
lbl.place(x=230, y=250)

txt=tk.Entry(root,width=20  ,bg="silver" ,fg="black",font=('times', 18, ' bold '),borderwidth=2,relief="solid")        
txt.place(x=500, y=250)

mess = tk.Label(root, text="" ,bg="silver"  ,fg="black"  ,width=17  ,height=0 , activebackground = "yellow" ,font=('times', 18, ' bold '),borderwidth=2,relief="solid")
mess.place(x=500, y=250)

bill = tk.Button(root,
                   text="Add to list",fg="black"  ,bg="#00CED1",
                    width=20 ,height=2,activebackground = "Red" ,font=('Comic Sans MS', 15, ' bold '),
                   command=bill)
bill.place(x=230,y=300)

display = tk.Button(root,
                   text="View BILL",fg="black"  ,bg="#00CED1",
                    width=20 ,height=2,activebackground = "Red" ,font=('Comic Sans MS', 15, ' bold '),
                   command=display)
display.place(x=500,y=300)

capture = tk.Button(root,
                   text="Generate Bill",fg="black"  ,bg="#00CED1",
                    width=41 ,height=2,activebackground = "Red" ,font=('Comic Sans MS', 15, ' bold '),
                   command=bill)
capture.place(x=240,y=400)

button = tk.Button(root, 
                   text="QUIT", fg="black"  ,bg="#FF0000"  ,
                   width=20 ,height=1,  activebackground = "Red" ,font=('Comic Sans MS', 15, ' bold '),
                   command=root.destroy)
button.place(x=375,y=500)



lbl = tk.Label(root, text="TEAM MEMBERS :  ",width=19  ,height=0  ,fg="blue"  ,bg="white" ,font=('Copperplate Gothic Bold', 15, ' bold ') )
lbl.place(x=570, y=560)

lbl = tk.Label(root, text="Harivyas S  ",width=30  ,height=0  ,fg="red"  ,bg="white" ,font=('Comic Sans MS', 15, ' bold ') )
lbl.place(x=590, y=585)
lbl = tk.Label(root, text="Abishek M",width=30  ,height=0  ,fg="red"  ,bg="white" ,font=('Comic Sans MS', 15, ' bold ') )
lbl.place(x=590, y=615)
lbl = tk.Label(root, text="Pavitharan M M ",width=30  ,height=0  ,fg="red"  ,bg="white" ,font=('Comic Sans MS', 15, ' bold ') )
lbl.place(x=590, y=644)


root.mainloop()

