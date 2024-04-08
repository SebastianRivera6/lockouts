from GenericDataModule import *
from LockoutBillingModule import *
import tkinter as tk
from tkinter import font
from PIL import Image, ImageTk
import subprocess


def option1():
    GenericDataDriver()
    completion_label.config(text="All new Lockouts Logged")
    show_image()



# Function to call ChargeDriver
def option2():
    ChargeDriver()
    completion_label.config(text="50 Lockouts Processed")
    show_image()


def show_image():
    # Load the PNG image from the media folder
    image_path = "media/WizardTuffy.png"
    image = Image.open(image_path)
    resized_image = image.resize((300, 300))  # Resize the image if necessary
    photo = ImageTk.PhotoImage(resized_image)

    # Update the image label to display the image
    image_label.config(image=photo)
    image_label.image = photo  # Keep a reference to avoid garbage collection

if __name__ == "__main__":
    # Create a main window
    window = tk.Tk()
    # Set font size for labels and buttons
    font_style = font.Font(family="Helvetica", size=14)

    # Set the window title
    window.title("CSUF Lockouts")

    # Set the window size
    window.geometry("800x800")

    # Load the background image
    background_image = Image.open("media/housing-splash.jpg")

    # Resize the background image to fit the window size
    background_image = background_image.resize((800, 800), Image.NEAREST)

    # Create a background image label
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = tk.Label(window, image=background_photo)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    # Create a label for the text with white color, bold font, and no text background
    label = tk.Label(window, text="CSUF Lockouts", font=(font_style, 24, "bold"), fg="white", bg="#00244E")
    label.grid(row=0, column=0, padx=10, pady=10, sticky="n")  # Align with the top of the window, starting at the same height as logo_label


    image = Image.open("media/Logo.png")
    resized_image = image.resize((100, 100))  # Resize the image if necessary
    logo_image = ImageTk.PhotoImage(resized_image)


    # Create a label for the image and place it in the top left corner
    logo_label = tk.Label(window, image=logo_image)
    logo_label.image = logo_image  # Keep a reference to avoid garbage collection
    logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")  # Place the label in the top left corner with padding


    # Add buttons
    button1 = tk.Button(window, text="Add to Generic Data", command=option1,font=(font_style, 18, "bold"), height=2, width=20)
    button1.grid(row=1, column=0, columnspan=1, pady=5, sticky="n")  # Add vertical padding to the first button

    button2 = tk.Button(window, text="Charge Lockouts", command=option2, font=(font_style, 18, "bold"), height=2, width=20)
    button2.grid(row=2, column=0, columnspan=1, pady=5, sticky="n")  # Add vertical padding to the second button

    # Label to display the image
    image_label = tk.Label(window)
    image_label.grid(row=3, column=0, columnspan=2, pady=5, sticky="n")  # Place the completion label at the bottom, spanning across two columns

    # Label to display completion message
    completion_label = tk.Label(window, text="", font=font_style, bg="#00244E", fg="white")
    completion_label.grid(row=4, column=0, columnspan=1, pady=5, sticky="n")  # Place the completion label at the bottom, spanning across two columns

    # Configure columns and rows to expand
    window.columnconfigure(0, weight=1)
    window.rowconfigure(0, weight=1)

    window.rowconfigure(1, weight=1)
    window.rowconfigure(2, weight=1)
    window.rowconfigure(3, weight=1)
    window.rowconfigure(4, weight=1)


    # Run the application
    window.mainloop()


