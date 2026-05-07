# evaluate.py
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
from sklearn.metrics import classification_report, confusion_matrix, ConfusionMatrixDisplay, accuracy_score, precision_score, recall_score, f1_score
import os
import json

def evaluate_and_plot(model, test_ds, history, class_names, output_dir=None, split_info=""):
    """
    Evaluate model and save results to disk
    
    Args:
        model: Trained model
        test_ds: Test dataset
        history: Training history
        class_names: List of class names
        output_dir: Directory to save results (if None, only print)
        split_info: String with split info (e.g., "split_60_40")
    """
    print("\n--- EWALUACJA MODELU ---")

    y_true = []
    for _, labels in test_ds:
        y_true.extend(labels.numpy())
    y_true = np.array(y_true)

    y_pred_probs = model.predict(test_ds)
    y_pred = np.argmax(y_pred_probs, axis=1) 

    # Calculate metrics
    accuracy = accuracy_score(y_true, y_pred)
    precision = precision_score(y_true, y_pred, average='weighted', zero_division=0)
    recall = recall_score(y_true, y_pred, average='weighted', zero_division=0)
    f1 = f1_score(y_true, y_pred, average='weighted', zero_division=0)
    
    report = classification_report(y_true, y_pred, target_names=class_names)
    
    print(f"\nRaport klasyfikacji ({split_info}):")
    print(report)
    print(f"\nPodsumowanie metryk:")
    print(f"  Dokładność (Accuracy):  {accuracy:.4f}")
    print(f"  Precyzja (Precision):   {precision:.4f}")
    print(f"  Czułość (Recall):       {recall:.4f}")
    print(f"  F1-Score:               {f1:.4f}")

    # Save results if output directory is provided
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        
        # 1. Save confusion matrix
        cm = confusion_matrix(y_true, y_pred)
        plt.figure(figsize=(10, 8)) # Zwiększyłem trochę rozmiar, żeby napisy się nie nakładały
        disp = ConfusionMatrixDisplay(confusion_matrix=cm, display_labels=class_names)
        disp.plot(cmap=plt.cm.Blues, ax=plt.gca(), xticks_rotation=45) # Używamy bieżącej osi
        plt.title(f'Macierz Pomyłek ({split_info})')
        plt.tight_layout()
        confusion_path = os.path.join(output_dir, 'confusion_matrix.png')
        plt.savefig(confusion_path, dpi=100, bbox_inches='tight', format='png')
        plt.close('all')
        if os.path.exists(confusion_path):
            print(f"✓ Macierz pomyłek zapisana: confusion_matrix.png ({os.path.getsize(confusion_path)} B)")
        else:
            print(f"✗ Błąd: Nie można zapisać macierzy pomyłek!")
        
        # 2. Save accuracy plot
        plt.figure(figsize=(10, 6))
        plt.plot(history.history['accuracy'], label='Dokładność treningowa', linewidth=2)
        plt.plot(history.history['val_accuracy'], label='Dokładność walidacyjna', linewidth=2)
        plt.xlabel('Epoka')
        plt.ylabel('Dokładność')
        plt.legend()
        plt.title(f'Krzywe uczenia ({split_info})')
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        accuracy_path = os.path.join(output_dir, 'accuracy_plot.png')
        plt.savefig(accuracy_path, dpi=100, bbox_inches='tight', format='png')
        plt.close('all')
        if os.path.exists(accuracy_path):
            print(f"✓ Wykres dokładności zapisany: accuracy_plot.png ({os.path.getsize(accuracy_path)} B)")
        else:
            print(f"✗ Błąd: Nie można zapisać wykresu dokładności!")
        
        # 3. Save metrics to JSON
        metrics_data = {
            'split_info': split_info,
            'accuracy': float(accuracy),
            'precision': float(precision),
            'recall': float(recall),
            'f1_score': float(f1),
            'confusion_matrix': cm.tolist(),
            'classification_report': report
        }
        with open(os.path.join(output_dir, 'metrics.json'), 'w') as f:
            json.dump(metrics_data, f, indent=2)
        print(f"✓ Metryki zapisane: metrics.json")
        
        # 4. Save classification report as text
        with open(os.path.join(output_dir, 'classification_report.txt'), 'w') as f:
            f.write(f"Split Info: {split_info}\n")
            f.write(f"{'='*50}\n\n")
            f.write(report)
            f.write(f"\n\nPodsumowanie metryk:\n")
            f.write(f"  Dokładność (Accuracy):  {accuracy:.4f}\n")
            f.write(f"  Precyzja (Precision):   {precision:.4f}\n")
            f.write(f"  Czułość (Recall):       {recall:.4f}\n")
            f.write(f"  F1-Score:               {f1:.4f}\n")
        print(f"✓ Raport tekstowy zapisany: classification_report.txt")
        
        return metrics_data
    
    return None