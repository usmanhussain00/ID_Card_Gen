import pandas as pd
from PIL import Image,ImageDraw,ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import tkinter as tk
from tkinter import filedialog,messagebox

FONT_PATH = 'arial.ttf'
#--------------------------------------------------------------------------------
#creating the id card
def create_id_card(candidate,font_path,columns):
    id_card_width = 400                 #This sets the id cards height to 400 pixels
    id_card_height = 600                  #This sets the id cards width to 600 pixel
    background_color = (255,255,255)     #(255,255,255)is the rgb value of the color white which is we are setting for the background color of the id card
    font_color = (0,0,0)                 #(0,0,0)is the rgb value of black is the font color of the id card

    id_card = Image.new('RGB',(id_card_width,id_card_height),background_color) #this creates a new image for a id carc to be drawn on with specified values
    draw = ImageDraw.Draw(id_card)                    #This creates a drawing object that allows us to draw a id card

    try:
        font = ImageFont.truetype(font_path,size=20)  #this loads the font file in the specified directory(font)
    except IOError:                                   #this line of code watches that if a error occurs like the specified file s not found then
        print(f"font file not found in the {font_path}")  #it will print this line of code
        return None

    y_position = 20   #specifies the y position were the text should start
    for column in columns:      #loops through each column name in the database
        if column != 'Picture Path'and pd.notna(candidate[column]):   # checks if the column is not picture path
            text = f"{column}:{candidate[column]}"   #creates a text string with column name and candidates corresponding value
            draw.text((20,y_position),text,font=font,fill=font_color)    #draws the text on id card from the specified position ,color ,font
            y_position += 40                                           #start the next value on the id card 40 coordinates + from the previous one

    picture_column = 'Picture Path'   #This defines that picture path in the columns hold the picture of the candidates
    if  picture_column  in columns and pd.notna(candidate[picture_column]):  #It checks if puicture path exist in the columns for the candidate and it is not null
        picture_path = candidate[picture_column]   #thsi code gets the picture path from the candidate and prints that it is being processed
        print(f"processing picture path:{picture_path}")
        if os.path.isfile(picture_path):   #it checks if the file given at the path exists
            try:
                picture = Image.open(picture_path)     #This opens the picture given at the picture path
                picture = picture.resize((200, 200))   #this resizes the picture
                id_card.paste(picture, (100, y_position))  #this paste the picture at the coordinates


            except Exception as e:
                print(f"ERROR: failed to open the picture file at {picture_path} with  error{e}")  #if while opening the image it catches a error it will print this image

        else:
            print(f"picture not found at {picture_path}")   #it gives a else statement that if the file s not found at the given path

    return id_card  #then thn function return the id card with picture pasted on to it

#---------------------------------------------------------------------------------
#function to generate pdf
def generate_pdf(id_cards,output_pdf):
    c = canvas.Canvas(output_pdf,pagesize=A4)   #canvas.Canvas is basically a blank page where you can draw things(output_pdf is the file path)and(A4 specify the size of pages in the output pdf)
    width,height = A4   #this line gets the height,width of the A4 page sizes and stores them into the variable called width,height

    for index, id_card in enumerate(id_cards):  #this function goes through every id card the enumerate function will give you the index and the image(id_card)of the id card
        id_card_path = f'temp_id_card{index}.jpg'  # this will create a temporary image file where  the current id card will be saved index function will make sure every file has a unique name
        id_card.save(id_card_path)   #this saves the current id card to the specified path(id_card_path)
        c.drawImage(id_card_path,0,0,width,height)  #this draws the id card to the current page of the pdf the image is placed at the (0,0) coordinates and is scaled to fit height,width
        c.showPage() #this ends the current page and makes the next id card on the next page each card is placed on different pages
        os.remove(id_card_path) #this will clean up the temporary image file in the system


    c.save()   #this will finalize and save the pdf to output_pdf

#--------------------------------------------------------------------------------
#reading the excel file
def main(excel_file,output_pdf):

    try:
        df = pd.read_excel(excel_file)   # this line of code is reading the excel file
    except FileNotFoundError:
        messagebox.showerror("Error",f"file not found at {excel_file}")          # this line of code will show a error messagebox if the file is not found
        return
    except Exception as e:
        messagebox.showerror("Error",f"can,t read the excel file: {str(e)}")     #this will show a error messagebox if the program can,t read the excel file
        return

    df = df.dropna(how='all',axis=1).dropna(how='all',axis=0)      #with we are removing all the empty column,row in the excel file so if the data is uneven or starting from the center it won,t affect the file because it will help it read those row,column which have data in it
    df.columns = df.iloc[0]             #this code will select the first row as the header for the column beneath it
    df = df[1:]                       #this select all the rows in the data frame except the first one (because of the slicing [1:])
    df.reset_index(drop=True,inplace=True)   #it will reset the index of the dataframe to 0 starting and repeating the same for the next line of code

    id_cards = []        #this is will create a empty list where all the id cards will be stored
    columns = df.columns  #With this function you are calling the names of all column
    print(f"column found in the excel file{columns}")
    for index, candidate in df.iterrows():  #this loop goes through every row and gives the index and data of the row
        print(f"processing candidate {index}: {candidate.to_dict()}")
        id_card = create_id_card(candidate,FONT_PATH,columns) #this function calls and generate a id card it usees details(canditates)and column names(column)and a specified font(FONT_PATH)
        if id_card:     #this checks if the id card was succesfully created  if yes
            id_cards.append(id_card)  #then this will add the id card to the list
        else:
            print(f"failed to create id card for the candidate{index}")   # or else it will print this line of code

    if id_cards:
        generate_pdf(id_cards,output_pdf) #this will check if  the id card were created and a output pdf was created
        messagebox.showinfo("success",f"pdf generated successfully: {output_pdf}") # then this line code will execute and show a message box which will current line of code
    else:
        messagebox.showwarning("ERROR",f"no id card were created") #else this line code will execute with this text printed on a message box

#---------------------------------------------------------------------------------

#This function is letting browse a excel file in your storage after the button is clicked and then saving it as pdf

def open_file():
    excel_file = filedialog.askopenfilename(title="open excel file",filetypes=(("Excel file","*.xlsx"),("All files","*.*")))
    if excel_file:
        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf",title="save pdf as",filetypes=(("PDF file","*.pdf"),("All files","*.*")))
        if output_pdf:
            main(excel_file , output_pdf)




#---------------------------------------------------------------------------------

#This setup is the for the button which helps you to the directory of the excel file

root  = tk.Tk()
root.title("ID card generator")

open_button = tk.Button(root,text='open Excel file',command=open_file)
open_button.pack(pady=20)

root.mainloop()


