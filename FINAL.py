import cv2
import os
import fitz  # PyMuPDF
import pytesseract
from docx import Document
import tkinter as tk
import tkinter.messagebox as messagebox
from tkinter import Label, Button
from tkinter import filedialog
from PIL import Image, ImageTk

# Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Camera

def capture_images(output_folder):
    cam = cv2.VideoCapture(1)
    cv2.namedWindow("NIMBUS")
    img_counter = 0

    while True:
        ret, frame = cam.read()
        if not ret:
            print("CAMERA FAILED")
            break
        cv2.imshow("APPLICATION", frame)

        k = cv2.waitKey(1)
        if k % 256 == 27:
            print("ESC hit, Closing the application")
            break
        elif k % 256 == 32:
            img_name = os.path.join(output_folder, f"Board_Image_{img_counter}.png")
            cv2.imwrite(img_name, frame)
            print(f"Image {img_name} saved!")
            img_counter += 1

    cam.release()
    cv2.destroyAllWindows()


# Image to PDF
def convert_images_to_pdf(image_folder, output_pdf_path):
    img_list = []
    for image_file in os.listdir(image_folder):
        if image_file.endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(image_folder, image_file)
            try:
                img = Image.open(image_path).convert('RGB')
                img_list.append(img)
            except Exception as e:
                print(f"Error processing image {image_file}: {e}")

    if img_list:
        img_list[0].save(output_pdf_path, save_all=True, append_images=img_list[1:])
        print(f"PDF successfully saved at {output_pdf_path}")
    else:
        print("No valid images found.")


# PDF to image
def pdf_to_images(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        images = []
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        return images
    except Exception as e:
        print(f"Error converting PDF to images: {e}")
        return []


# OCR
def ocr_handwritten_text(image):
    try:
        config = '--psm 6 --oem 1'  # Block of text, LSTM engine for handwriting
        text = pytesseract.image_to_string(image, lang='eng', config=config)
        return text
    except Exception as e:
        print(f"Error during OCR: {e}")
        return ""


# Final World Doc
def save_to_word(text, output_word_path):
    try:
        doc = Document()
        doc.add_paragraph(text)
        doc.save(output_word_path)
        print(f"Text successfully saved to {output_word_path}")
    except Exception as e:
        print(f"Error saving to Word: {e}")


def detect_handwritten_text_to_word(input_path, output_word_path):
    if input_path.endswith('.pdf'):
        images = pdf_to_images(input_path)
    else:
        images = [Image.open(os.path.join(input_path, f)) for f in os.listdir(input_path)
                  if f.endswith(('.png', '.jpg', '.jpeg'))]

    if not images:
        print("No images found. Exiting.")
        return

    all_text = ""
    for i, image in enumerate(images):
        print(f"Processing page {i + 1}")
        text = ocr_handwritten_text(image)
        all_text += text + "\n\n"

    save_to_word(all_text, output_word_path)


def open_folder():
    folder = filedialog.askdirectory()
    folder_label.config(text=folder)


def save_file():
    file = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")],
        title="Select Output Word Document"
    )
    if file:
        save_label.config(text=file)


def start_conversion():
    output_folder = folder_label.cget("text")  # Get the folder selected by the user
    output_word_path = save_label.cget("text")  # Get the output file path selected by the user

    if output_folder == "No folder selected" or output_word_path == "No file selected":
        messagebox.showwarning("Input Error", "Please select a folder and output file.")
        return

    pdf_output = os.path.join(output_folder, "converted.pdf")

    try:
        capture_images(output_folder)
        convert_images_to_pdf(output_folder, pdf_output)
        detect_handwritten_text_to_word(pdf_output, output_word_path)
        messagebox.showinfo("Success", "OCR and conversion process completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Main Application Window
root = tk.Tk()
root.title("OCR Application - NIMBUS")
root.geometry("640x480")  # Set window size

# Load background image
bg_image = Image.open("C:\\Users\\Raj\\Downloads\\NIMBUS-bg.png")  # Use your own image path here
bg_image = bg_image.resize((640, 480), Image.Resampling.LANCZOS)  # Resize to fit window
bg_photo = ImageTk.PhotoImage(bg_image)

# Set background
bg_label = Label(root, image=bg_photo)
bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# Adjust grid to center widgets and push them down
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=1)
root.grid_rowconfigure(4, weight=1)
root.grid_rowconfigure(5, weight=1)
root.grid_rowconfigure(6, weight=1)
root.grid_rowconfigure(7, weight=1)
root.grid_columnconfigure(0, weight=1)

# Nimbus Heading
nimbus_label = Label(root, text="NIMBUS", bg="white", font=("Impact", 24), fg="black")
nimbus_label.grid(row=0, column=0, padx=260, pady=20, sticky="nsew", columnspan=2)

# Instructions
instr_label = Label(root, text="Step 1: Select Folder to Save Captured Images", bg="white", font=("Impact", 12))
instr_label.grid(row=1, column=0, padx=160, pady=10, sticky="nsew", columnspan=2)

# Folder Selection Button
folder_button = Button(root, text="Select Folder", command=open_folder)
folder_button.grid(row=2, column=0, padx=260, pady=10, sticky="nsew", columnspan=2)

# Label to Display Selected Folder
folder_label = Label(root, text="No folder selected", bg="white", font=("Arial", 10))
folder_label.grid(row=3, column=0, padx=260, pady=10, sticky="nsew", columnspan=2)

# Instructions for Saving
save_instr_label = Label(root, text="Step 2: Select Output Word Document", bg="white", font=("Impact", 12))
save_instr_label.grid(row=4, column=0, padx=190, pady=10, sticky="nsew", columnspan=2)

# Save File Button
save_button = Button(root, text="Select Save Location", command=save_file)
save_button.grid(row=5, column=0, padx=240, pady=10, sticky="nsew", columnspan=2)

# Label to Display Save Location
save_label = Label(root, text="No file selected", bg="white", font=("Arial", 10))
save_label.grid(row=6, column=0, padx=260, pady=10, sticky="nsew", columnspan=2)

# Start Conversion Button
convert_button = Button(root, text="Capture Images and Start OCR", command=start_conversion)
convert_button.grid(row=7, column=0, padx=190, pady=20, sticky="nsew", columnspan=2)

root.mainloop()
