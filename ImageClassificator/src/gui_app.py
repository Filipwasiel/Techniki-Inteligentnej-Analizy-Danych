import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import tensorflow as tf
import numpy as np
import threading
import os

# Importy z Twojej struktury projektu
from src.models_factory import create_model
from src.data_loader import load_data
from src.data_manager import initialize_raw_data, split_data, SELECTED_CLASSES
from src import config

class FoodClassifierGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Ustawienia okna
        self.title("Klasyfikator Jedzenia ")
        self.geometry("1200x900")
        ctk.set_appearance_mode("dark")
        
        # Stan aplikacji
        self.model_wrapper = None
        self.class_names = SELECTED_CLASSES
        self.is_training = False
        self.bar_widgets = {}

        # --- UKŁAD (GRID) ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 1. PANEL BOCZNY (Sidebar)
        self.sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        ctk.CTkLabel(self.sidebar, text="USTAWIENIA", font=("Arial", 22, "bold")).pack(pady=25)

        self.model_choice = ctk.CTkOptionMenu(self.sidebar, values=["resnet", "mobilenet", "simple_cnn"])
        self.model_choice.pack(pady=10, padx=25, fill="x")

        # Sekcja suwaka z dynamiczną etykietą
        ctk.CTkLabel(self.sidebar, text="Zbiór treningowy (%):").pack(anchor="w", padx=25)
        self.split_info_label = ctk.CTkLabel(self.sidebar, text="Wybrano: 70%", font=("Arial", 12, "italic"), text_color="cyan")
        self.split_info_label.pack(anchor="w", padx=25)
        
        self.split_slider = ctk.CTkSlider(
            self.sidebar, 
            from_=0.1, 
            to=0.9, 
            number_of_steps=8, 
            command=self.update_slider_label
        )
        self.split_slider.set(0.7)
        self.split_slider.pack(pady=(5, 10), padx=25, fill="x")

        self.epochs_entry = ctk.CTkEntry(self.sidebar)
        self.epochs_entry.insert(0, "15")
        self.epochs_entry.pack(pady=5, padx=25, fill="x")

        self.batch_entry = ctk.CTkEntry(self.sidebar)
        self.batch_entry.insert(0, "32")
        self.batch_entry.pack(pady=5, padx=25, fill="x")

        # Niebieski przycisk treningu
        self.btn_train = ctk.CTkButton(
            self.sidebar, 
            text="ROZPOCZNIJ NAUKĘ", 
            command=self.train_model_thread, 
            fg_color="#1f538d", 
            hover_color="#14375e"
        )
        self.btn_train.pack(pady=20, padx=25, fill="x")

        self.status_label = ctk.CTkLabel(self.sidebar, text="Status: Gotowy", text_color="gray")
        self.status_label.pack(side="bottom", pady=30)

        # 2. PANEL GŁÓWNY
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1) # Centrowanie wewnątrz

        ctk.CTkLabel(self.main_frame, text="TESTOWANIE I ANALIZA", font=("Arial", 24, "bold")).pack(pady=20, anchor="center")

        # PASEK POSTĘPU TRENINGU (Ukryty na starcie)
        self.training_info_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        
        self.prog_bar = ctk.CTkProgressBar(self.training_info_frame, width=400)
        self.prog_bar.set(0)
        self.prog_bar.pack(side="left", padx=10)
        
        self.prog_label = ctk.CTkLabel(self.training_info_frame, text="0/0 | Acc: 0.00", font=("Arial", 14))
        self.prog_label.pack(side="left", padx=10)

        self.btn_upload = ctk.CTkButton(self.main_frame, text="Wybierz zdjęcie z dysku", command=self.upload_and_predict, height=45)
        self.btn_upload.pack(pady=10, anchor="center")

        self.image_display = ctk.CTkLabel(self.main_frame, text="Podgląd zdjęcia", width=400, height=400, fg_color="#1e1e1e", corner_radius=15)
        self.image_display.pack(pady=15, anchor="center")

        # PANEL WYNIKÓW / WYKRES (Ukryty na starcie)
        self.res_frame = ctk.CTkFrame(self.main_frame, fg_color="#2b2b2b")
        
        ctk.CTkLabel(self.res_frame, text="ROZKŁAD PRAWDOPODOBIEŃSTWA", font=("Arial", 16, "bold")).pack(pady=10, anchor="center")

        for food_class in SELECTED_CLASSES:
            row = ctk.CTkFrame(self.res_frame, fg_color="transparent")
            row.pack(fill="x", padx=20, pady=3)
            
            lbl = ctk.CTkLabel(row, text=f"{food_class.replace('_', ' ').title():18}", font=("Courier", 13))
            lbl.pack(side="left")
            
            progress = ctk.CTkProgressBar(row, width=300, height=12)
            progress.set(0)
            progress.pack(side="left", padx=15)
            
            pct_lbl = ctk.CTkLabel(row, text="0.0%", font=("Arial", 12), width=50)
            pct_lbl.pack(side="left")
            
            self.bar_widgets[food_class] = (progress, pct_lbl)

    # --- LOGIKA ---

    def update_slider_label(self, value):
        self.split_info_label.configure(text=f"Wybrano: {int(value * 100)}%")

    def log(self, text, color="white"):
        self.status_label.configure(text=f"Status: {text}", text_color=color)

    def train_model_thread(self):
        if self.is_training: return
        threading.Thread(target=self.do_auto_train, daemon=True).start()

    def do_auto_train(self):
        self.is_training = True
        self.btn_train.configure(state="disabled")
        
        # POKAZUJEMY pasek, UKRYWAMY stare wyniki
        self.res_frame.pack_forget()
        self.training_info_frame.pack(pady=10, anchor="center")
        self.prog_bar.set(0)
        self.prog_label.configure(text="Przygotowanie danych...")
        
        try:
            # 1. Automatyczne przygotowanie danych
            self.log("Przetwarzanie plików...", "yellow")
            initialize_raw_data()
            split_val = self.split_slider.get()
            split_data(train_split=split_val)

            # 2. Konfiguracja
            epochs = int(self.epochs_entry.get())
            config.BATCH_SIZE = int(self.batch_entry.get())
            model_name = self.model_choice.get()

            train_ds, test_ds, classes = load_data()
            self.class_names = classes
            
            self.model_wrapper = create_model(model_name, num_classes=len(classes))
            self.model_wrapper.build()

            self.log(f"Nauka: {model_name}", "orange")
            
            # 3. Pętla treningowa z aktualizacją GUI
            for epoch in range(epochs):
                history = self.model_wrapper.model.fit(
                    train_ds, 
                    epochs=1, 
                    validation_data=test_ds, 
                    verbose=0
                )
                
                acc = history.history['accuracy'][0]
                current = epoch + 1
                progress_val = current / epochs
                
                self.prog_bar.set(progress_val)
                self.prog_label.configure(text=f"{current}/{epochs} | Acc: {acc:.4f}")
                self.update_idletasks()

            self.log("Gotowy!", "#28a745")
            messagebox.showinfo("Sukces", "Model został wytrenowany pomyślnie!")
            
        except Exception as e:
            self.log("Błąd!", "red")
            messagebox.showerror("Błąd", str(e))
        finally:
            self.is_training = False
            self.btn_train.configure(state="normal")
            self.training_info_frame.pack_forget()

    def upload_and_predict(self):
        if not self.model_wrapper:
            messagebox.showwarning("Brak modelu", "Najpierw przeprowadź naukę modelu!")
            return

        path = filedialog.askopenfilename(filetypes=[("Zdjęcia", "*.jpg *.jpeg *.png")])
        if not path: return

        # Podgląd obrazu
        pil_img = Image.open(path)
        img_ctk = ctk.CTkImage(light_image=pil_img, size=(400, 400))
        self.image_display.configure(image=img_ctk, text="")

        # Przygotowanie do TF
        img_raw = tf.keras.utils.load_img(path, target_size=config.IMG_SIZE)
        img_array = tf.keras.utils.img_to_array(img_raw)
        img_array = tf.expand_dims(img_array, 0)

        # Predykcja i pokazanie wyników
        self.log("Analiza...", "cyan")
        preds = self.model_wrapper.model.predict(img_array)
        scores = tf.nn.softmax(preds[0]).numpy()
        max_idx = np.argmax(scores)

        # POKAZUJEMY ramkę z wykresami
        self.res_frame.pack(pady=20, padx=40, fill="none", anchor="center")

        for i, food_class in enumerate(self.class_names):
            confidence = scores[i]
            progress_bar, pct_label = self.bar_widgets[food_class]
            
            progress_bar.set(confidence)
            pct_label.configure(text=f"{confidence*100:.1f}%")
            
            if i == max_idx:
                progress_bar.configure(progress_color="#28a745")
                pct_label.configure(text_color="#28a745")
            else:
                progress_bar.configure(progress_color="#3b8ed0")
                pct_label.configure(text_color="white")

        self.log("Zakończono", "gray")

if __name__ == "__main__":
    app = FoodClassifierGUI()
    app.mainloop()