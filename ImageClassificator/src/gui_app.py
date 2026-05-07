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
from src.evaluate import evaluate_and_plot
from src import config

class HistoryWrapper:
    def __init__(self, history_dict):
        self.history = history_dict

class FoodClassifierGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Food Analyzer - Dashboard")
        self.geometry("1400x950")
        ctk.set_appearance_mode("dark")
        
        self.model_wrapper = None
        self.class_names = SELECTED_CLASSES
        self.is_training = False
        self.bar_widgets = {}

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 1. PANEL BOCZNY
        self.sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        ctk.CTkLabel(self.sidebar, text="USTAWIENIA", font=("Arial", 22, "bold")).pack(pady=25)
        
        self.model_choice = ctk.CTkOptionMenu(self.sidebar, values=["resnet", "mobilenet", "simple_cnn"])
        self.model_choice.pack(pady=10, padx=25, fill="x")

        ctk.CTkLabel(self.sidebar, text="Zbiór treningowy (%):").pack(anchor="w", padx=25)
        self.split_info_label = ctk.CTkLabel(self.sidebar, text="Wybrano: 70%", font=("Arial", 12, "italic"), text_color="cyan")
        self.split_info_label.pack(anchor="w", padx=25)
        
        self.split_slider = ctk.CTkSlider(self.sidebar, from_=0.1, to=0.9, number_of_steps=8, command=self.update_slider_label)
        self.split_slider.set(0.7)
        self.split_slider.pack(pady=(5, 10), padx=25, fill="x")

        self.epochs_entry = ctk.CTkEntry(self.sidebar)
        self.epochs_entry.insert(0, "15")
        self.epochs_entry.pack(pady=5, padx=25, fill="x")

        self.btn_train = ctk.CTkButton(self.sidebar, text="ROZPOCZNIJ NAUKĘ", command=self.train_model_thread, fg_color="#1f538d")
        self.btn_train.pack(pady=20, padx=25, fill="x")

        self.status_label = ctk.CTkLabel(self.sidebar, text="Status: Gotowy", text_color="gray")
        self.status_label.pack(side="bottom", pady=30)

        # 2. SYSTEM ZAKŁADEK
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        self.tab_predict = self.tabview.add("Predykcja")
        self.tab_stats = self.tabview.add("Statystyki")

        self._setup_predict_tab()
        self._setup_stats_tab()

    def _setup_predict_tab(self):
        container = ctk.CTkScrollableFrame(self.tab_predict, fg_color="transparent")
        container.pack(fill="both", expand=True)
        container.grid_columnconfigure(0, weight=1)

        self.training_info_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.prog_bar = ctk.CTkProgressBar(self.training_info_frame, width=400)
        self.prog_bar.pack(side="left", padx=10)
        self.prog_label = ctk.CTkLabel(self.training_info_frame, text="0/0 | Acc: 0.00")
        self.prog_label.pack(side="left", padx=10)

        # NOWOŚĆ: Główny napis z wynikiem klasyfikacji
        self.result_main_label = ctk.CTkLabel(container, text="Wgraj zdjęcie, aby rozpocząć", font=("Arial", 20, "bold"), text_color="cyan")
        self.result_main_label.pack(pady=10)

        self.btn_upload = ctk.CTkButton(container, text="Wybierz zdjęcie", command=self.upload_and_predict, height=40)
        self.btn_upload.pack(pady=15)

        self.image_display = ctk.CTkLabel(container, text="Brak obrazu", width=400, height=400, fg_color="#1e1e1e", corner_radius=15)
        self.image_display.pack(pady=10)

        # Podpis pomocniczy nad paskami
        self.confidence_title = ctk.CTkLabel(container, text="Pewność modelu dla poszczególnych klas:", font=("Arial", 14, "italic"))
        # Nie pakujemy go tutaj, pojawi się razem z res_frame

        self.res_frame = ctk.CTkFrame(container, fg_color="#2b2b2b")
        for food_class in SELECTED_CLASSES:
            row = ctk.CTkFrame(self.res_frame, fg_color="transparent")
            row.pack(fill="x", padx=20, pady=2)
            lbl = ctk.CTkLabel(row, text=f"{food_class.replace('_', ' ').title():18}", font=("Courier", 13))
            lbl.pack(side="left")
            
            # Ustawienie domyślnego koloru na niebieski (#1f538d)
            pb = ctk.CTkProgressBar(row, width=300, progress_color="#1f538d")
            pb.set(0)
            pb.pack(side="left", padx=15)
            
            pct = ctk.CTkLabel(row, text="0.0%", width=50)
            pct.pack(side="left")
            self.bar_widgets[food_class] = (pb, pct)

    def _setup_stats_tab(self):
        # Główny kontener bez scrolla, żeby łatwiej kontrolować układ "jednej strony"
        self.stats_main_container = ctk.CTkFrame(self.tab_stats, fg_color="transparent")
        self.stats_main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Sekcja na WYKRESY (obok siebie)
        self.plots_row = ctk.CTkFrame(self.stats_main_container, fg_color="transparent")
        self.plots_row.pack(fill="x", pady=10)

        # Sekcja na TABELĘ METRYK
        self.table_frame = ctk.CTkFrame(self.stats_main_container, fg_color="#2b2b2b")
        self.table_frame.pack(fill="both", expand=True, pady=10)
        
        self.no_data_lbl = ctk.CTkLabel(self.table_frame, text="Dane pojawią się po treningu")
        self.no_data_lbl.pack(pady=50)

    def do_auto_train(self):
        self.is_training = True
        self.btn_train.configure(state="disabled")
        self.training_info_frame.pack(pady=10)
        
        try:
            self.log("Przygotowanie danych...", "yellow")
            initialize_raw_data()
            split_val = self.split_slider.get()
            split_data(train_split=split_val)

            epochs = int(self.epochs_entry.get())
            train_ds, test_ds, classes = load_data()
            self.class_names = classes
            
            self.model_wrapper = create_model(self.model_choice.get(), num_classes=len(classes))
            self.model_wrapper.build()

            self.log("Trenowanie...", "orange")
            hist_dict = {'accuracy': [], 'val_accuracy': []}

            for epoch in range(epochs):
                h = self.model_wrapper.model.fit(train_ds, epochs=1, validation_data=test_ds, verbose=0)
                # Zbieramy dane do słownika ręcznie
                hist_dict['accuracy'].append(h.history['accuracy'][0])
                hist_dict['val_accuracy'].append(h.history['val_accuracy'][0])
                
                self.prog_bar.set((epoch + 1) / epochs)
                self.prog_label.configure(text=f"{epoch+1}/{epochs} | Acc: {h.history['accuracy'][0]:.4f}")
                self.update_idletasks()

            self.log("Ewaluacja...", "cyan")
            output_dir = "results"
            if not os.path.exists(output_dir): os.makedirs(output_dir)
            
            # --- ROZWIĄZANIE BŁĘDU .history ---
            class KerasHistoryObject:
                def __init__(self, d):
                    self.history = d
            
            # Pakujemy słownik w obiekt, który udaje wynik z model.fit()
            wrapped_history = KerasHistoryObject(hist_dict)
            
            # Wywołanie Twojej funkcji z evaluate.py
            evaluate_and_plot(
                model=self.model_wrapper.model, 
                test_ds=test_ds, 
                history=wrapped_history, 
                class_names=self.class_names, 
                output_dir=output_dir, 
                split_info=f"split_{int(split_val*100)}"
            )

            # Odświeżenie widoku statystyk
            self.display_stats(output_dir)
            self.log("Gotowy!", "#28a745")
            messagebox.showinfo("Sukces", "Trening i raporty zakończone!")
            
        except Exception as e:
            self.log("Błąd!", "red")
            print(f"DEBUG: {e}") # Sprawdź konsolę w razie problemów
            messagebox.showerror("Błąd", str(e))
        finally:
            self.is_training = False
            self.btn_train.configure(state="normal")
            self.training_info_frame.pack_forget()

    def display_stats(self, folder):
        # 1. Czyszczenie starych widżetów
        for widget in self.plots_row.winfo_children(): widget.destroy()
        for widget in self.table_frame.winfo_children(): widget.destroy()
        
        if hasattr(self, 'no_data_lbl') and self.no_data_lbl.winfo_exists():
            self.no_data_lbl.destroy()

        # 2. Wyświetlanie obrazów OBOK SIEBIE (50% szerokości każdy)
        acc_path = os.path.join(folder, "accuracy_plot.png")
        cm_path = os.path.join(folder, "confusion_matrix.png")

        # Rozmiar dobrany pod okno 1400px (50% szerokości z marginesami)
        target_width = 680 
        target_height = 500

        if os.path.exists(acc_path):
            img_acc = ctk.CTkImage(light_image=Image.open(acc_path), size=(target_width, target_height))
            lbl_acc = ctk.CTkLabel(self.plots_row, image=img_acc, text="")
            lbl_acc.pack(side="left", expand=True, fill="both", padx=5)

        if os.path.exists(cm_path):
            img_cm = ctk.CTkImage(light_image=Image.open(cm_path), size=(target_width, target_height))
            lbl_cm = ctk.CTkLabel(self.plots_row, image=img_cm, text="")
            lbl_cm.pack(side="left", expand=True, fill="both", padx=5)

        # 3. Budowanie TABELI metryk
        report_path = os.path.join(folder, "classification_report.txt")
        if os.path.exists(report_path):
            with open(report_path, "r") as f:
                lines = f.readlines()
            
            headers = ["KLASA", "PRECISION", "RECALL", "F1-SCORE", "SUPPORT"]
            for col_idx, header in enumerate(headers):
                h_lbl = ctk.CTkLabel(self.table_frame, text=header, font=("Arial", 13, "bold"), text_color="cyan")
                h_lbl.grid(row=0, column=col_idx, padx=25, pady=10, sticky="nsew")

            row_idx = 1
            for line in lines:
                parts = line.split()
                # Dopasowanie nazwy klasy (zamiana spacji na podkreślniki jeśli trzeba)
                valid_classes = [c.replace(' ', '_') for c in self.class_names]
                
                if len(parts) >= 5 and parts[0] in valid_classes:
                    for col_idx, val in enumerate(parts[:5]):
                        c_lbl = ctk.CTkLabel(self.table_frame, text=val, font=("Arial", 12))
                        c_lbl.grid(row=row_idx, column=col_idx, padx=25, pady=2)
                    row_idx += 1
            
            # Wyświetlenie końcowej dokładności
            for line in lines:
                if "accuracy" in line and len(line.split()) >= 2:
                    acc_val = line.split()[-2]
                    acc_lbl = ctk.CTkLabel(self.table_frame, text=f"OVERALL ACCURACY: {acc_val}", 
                                           font=("Arial", 13, "bold"), text_color="#28a745")
                    acc_lbl.grid(row=row_idx+1, column=0, columnspan=5, pady=15)

    def update_slider_label(self, value):
        self.split_info_label.configure(text=f"Wybrano: {int(value * 100)}%")

    def log(self, text, color="white"):
        self.status_label.configure(text=f"Status: {text}", text_color=color)

    def train_model_thread(self):
        if self.is_training: return
        threading.Thread(target=self.do_auto_train, daemon=True).start()

    

    def upload_and_predict(self):
            if not self.model_wrapper:
                messagebox.showwarning("Uwaga", "Wytrenuj model!")
                return
            path = filedialog.askopenfilename(filetypes=[("Zdjęcia", "*.jpg *.jpeg *.png")])
            if not path: return

            pil_img = Image.open(path)
            img_ctk = ctk.CTkImage(light_image=pil_img, size=(400, 400))
            self.image_display.configure(image=img_ctk, text="")

            img_raw = tf.keras.utils.load_img(path, target_size=config.IMG_SIZE)
            img_array = tf.keras.utils.img_to_array(img_raw)
            img_array = tf.expand_dims(img_array, 0)

            preds = self.model_wrapper.model.predict(img_array)
            scores = tf.nn.softmax(preds[0]).numpy()
            
            # Wyznaczamy indeks klasy z najwyższym prawdopodobieństwem
            max_idx = np.argmax(scores)
            predicted_name = self.class_names[max_idx].replace('_', ' ').upper()

            # Aktualizacja głównego napisu o wykrytej klasie
            self.result_main_label.configure(text=f"WYKRYTO: {predicted_name}", text_color="#28a745")

            # Pokazujemy nagłówek i ramkę z wynikami
            self.confidence_title.pack(pady=5)
            self.res_frame.pack(pady=20, fill="none", anchor="center")

            # Aktualizacja pasków i ich kolorów
            for i, food_class in enumerate(self.class_names):
                confidence = scores[i]
                pb, pct = self.bar_widgets[food_class]
                
                pb.set(confidence)
                pct.configure(text=f"{confidence*100:.1f}%")
                
                # Jeśli to klasa z najwyższym wynikiem -> zielony, w przeciwnym razie -> niebieski
                if i == max_idx:
                    pb.configure(progress_color="#28a745")
                    pct.configure(text_color="#28a745")
                else:
                    pb.configure(progress_color="#1f538d")
                    pct.configure(text_color="white")

if __name__ == "__main__":
    app = FoodClassifierGUI()
    app.mainloop()