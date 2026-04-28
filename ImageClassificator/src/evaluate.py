# evaluate.py
import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import classification_report, confusion_matrix, ConfusionMatrixDisplay

def evaluate_and_plot(model, test_ds, history):
    print("\n--- EWALUACJA MODELU ---")

    y_true = []
    for images, labels in test_ds:
        y_true.extend(labels.numpy())
    y_true = np.array(y_true)

    y_pred_probs = model.predict(test_ds)
    y_pred = (y_pred_probs > 0.5).astype(int)

    print("\nRaport klasyfikacji (Dokładność, Precyzja, Czułość, F1):")
    print(classification_report(y_true, y_pred, target_names=['Cats', 'Dogs']))

    cm = confusion_matrix(y_true, y_pred)
    disp = ConfusionMatrixDisplay(confusion_matrix=cm, display_labels=['Cats', 'Dogs'])
    disp.plot(cmap=plt.cm.Blues)
    plt.title('Macierz Pomyłek')
    plt.show()

    plt.plot(history.history['accuracy'], label='Dokładność treningowa')
    plt.plot(history.history['val_accuracy'], label='Dokładność walidacyjna')
    plt.xlabel('Epoka')
    plt.ylabel('Dokładność')
    plt.legend()
    plt.title('Krzywe uczenia')
    plt.show()