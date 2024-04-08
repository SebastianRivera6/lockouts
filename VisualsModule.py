from GenericDataModule import *
from LockoutBillingModule import *
import tkinter as tk
from tkinter import font
from PIL import Image, ImageTk
import subprocess

def option1(window, font_style):
    GenericDataDriver()
    completion_label = tk.Label(window, text="", font=font_style, bg="#00244E", fg="white")
    completion_label.grid(row=4, column=0, columnspan=1, pady=5,sticky="n")
    completion_label.config(text="All new Lockouts Logged")
    show_image(window)


# Function to call ChargeDriver
def option2(window, font_style):
    ChargeDriver()
    # Label to display completion message
    completion_label = tk.Label(window, text="", font=font_style, bg="#00244E", fg="white")
    completion_label.grid(row=4, column=0, columnspan=1, pady=5,sticky="n")
    completion_label.config(text="50 Lockouts Processed")
    show_image(window)


def show_image(window):
    # Load the PNG image from the media folder
    # Label to display the image
    image_label = tk.Label(window)
    image_label.grid(row=3, column=0, columnspan=2, pady=5, sticky="n")
    image_path = "media/WizardTuffy.png"
    image = Image.open(image_path)
    resized_image = image.resize((300, 300))
    photo = ImageTk.PhotoImage(resized_image)
    
    # Update the image label to display the image
    image_label.config(image=photo)
    image_label.image = photo

def WindowLoop(window, background_image, institution_logo, font_style):
    
    # Set the window title
    window.title("CSUF Lockouts")

    # Set the window size
    window.geometry("800x800")

    # Load the background image
    background_image = Image.open(background_image)

    # Resize the background image to fit the window size
    background_image = background_image.resize((800, 800), Image.NEAREST)

    # Create a background image label
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = tk.Label(window, image=background_photo)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    # Create a label for the text with white color, bold font, and no text background
    label = tk.Label(window, text="CSUF Lockouts", font=(font_style, 24, "bold"), fg="white", bg="#00244E")
    label.grid(row=0, column=0, padx=10, pady=10, sticky="n")

    image = Image.open(institution_logo)
    resized_image = image.resize((100, 100))
    logo_image = ImageTk.PhotoImage(resized_image)

    # Create a label for the image and place it in the top left corner
    logo_label = tk.Label(window, image=logo_image)
    logo_label.image = logo_image
    logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

    # Add buttons
    button1 = tk.Button(window, text="Add to Generic Data", command=lambda: option1(window, font_style), font=font_style, height=2, width=20)
    button1.grid(row=1, column=0, columnspan=1, pady=5, sticky="n")

    button2 = tk.Button(window, text="Charge Lockouts", command=lambda: option2(window, font_style), font=font_style, height=2, width=20)
    button2.grid(row=2, column=0, columnspan=1, pady=5, sticky="n")

    # Configure columns and rows to expand
    window.columnconfigure(0, weight=1)
    window.rowconfigure(0, weight=1)
    window.rowconfigure(1, weight=1)
    window.rowconfigure(2, weight=1)
    window.rowconfigure(3, weight=1)
    window.rowconfigure(4, weight=1)

    # Run the application
    window.mainloop()
