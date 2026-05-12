# main.py
import logging
import os
import sys
import datetime
from pathlib import Path
from xml.parsers.expat import model

import tensorflow as tf

from src import config, data_manager
from src.data_loader import load_data
from src.models_factory import create_model
from src.evaluate import evaluate_and_plot
from src.data_manager import SELECTED_CLASSES

# Suppress TensorFlow logging
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
tf.get_logger().setLevel(logging.ERROR)


def run_single_experiment(model_name: str, train_split: float, output_dir: str):
    """
    Run single experiment with given model and train/test split
    
    Args:
        model_name: Name of model to use
        train_split: Training data percentage (0.0-1.0)
        output_dir: Directory to save results
    
    Returns:
        Dictionary with experiment results
    """
    test_split = 1.0 - train_split
    print(f"\n{'='*70}")
    print(f"EKSPERYMENT: Model={model_name} | Split: {train_split*100:.0f}% train / {test_split*100:.0f}% test")
    print(f"{'='*70}\n")
    
    try:
        # 1. Prepare data
        print("1. Przygotowanie danych...")
        data_manager.initialize_raw_data()
        data_manager.split_data(train_split=train_split)
        
        # 2. Load data
        print("2. Ładowanie danych z dysku...")
        train_ds, test_ds, class_names = load_data()
        num_detected_classes = len(class_names)
        print(f"Wykryto klas: {num_detected_classes}")

        # 3. Create and build model
        print("3. Budowanie struktury modelu...")
        input_shape = (config.IMG_SIZE[0], config.IMG_SIZE[1], 3)
        model = create_model(model_name, input_shape, num_classes=num_detected_classes)
        model.build()
        
        # 4. Train model
        print("4. Rozpoczynanie trenowania...")
        history = model.train(
            train_ds=train_ds,
            epochs=config.EPOCHS,
            validation_data=test_ds
        )
        
        # 5. Evaluate and save results
        print("5. Generowanie raportów...")
        os.makedirs(output_dir, exist_ok=True)
        split_info = f"{model_name}_split_{train_split*100:.0f}_{test_split*100:.0f}"
        metrics = evaluate_and_plot(
            model.model,
            test_ds,
            history,
            class_names=class_names,
            output_dir=output_dir,
            split_info=split_info
        )
        
        print(f"\n✓ Eksperyment zakończony pomyślnie!")
        print(f"  Model: {model_name}")
        print(f"  Podział: {train_split*100:.0f}% train / {test_split*100:.0f}% test")
        print(f"  Wyniki: {output_dir}\n")
        
        return {
            'status': 'SUCCESS',
            'model': model_name,
            'split': f"{train_split*100:.0f}_{test_split*100:.0f}",
            'output_dir': output_dir,
            'metrics': metrics
        }
        
    except Exception as e:
        print(f"\n✗ Błąd w eksperymencie: {str(e)}")
        import traceback
        traceback.print_exc()
        return {
            'status': 'FAILED',
            'model': model_name,
            'split': f"{train_split*100:.0f}_{test_split*100:.0f}",
            'error': str(e)
        }


def run_experiments(
    model_names=None,
    train_splits=None,
    results_dir='results'
):
    """
    Run series of experiments with different models and splits
    
    Args:
        model_names: List of model names to test (default: ['simple_cnn'])
        train_splits: List of train/test splits as train percentages (default: [0.6, 0.7, 0.8, 0.9])
        results_dir: Base directory for results (creates timestamped subdirectory)
    """
    # Set defaults
    if model_names is None:
        model_names = ['simple_cnn']
    if train_splits is None:
        train_splits = [0.6, 0.7, 0.8, 0.9]
    
    # Create results directory with timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    results_base = os.path.join(results_dir, f"experiments_{timestamp}")
    os.makedirs(results_base, exist_ok=True)
    
    # Print header
    print("\n" + "="*70)
    print("SERIA EKSPERYMENTÓW - KLASYFIKACJA OBRAZÓW")
    print("="*70)
    print(f"\nCzas startu: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Folder wyników: {results_base}")
    print(f"Modele: {', '.join(model_names)}")
    print(f"Podziały danych: {[f'{s*100:.0f}% train' for s in train_splits]}")
    print(f"Razem eksperymentów: {len(model_names) * len(train_splits)}\n")
    
    results_summary = []
    total_experiments = len(model_names) * len(train_splits)
    current_experiment = 0
    
    # Run all combinations
    for model_name in model_names:
        for train_split in train_splits:
            current_experiment += 1
            
            # Create subdirectory for this experiment
            split_name = f"{model_name}_split_{train_split*100:.0f}_{(1-train_split)*100:.0f}"
            experiment_dir = os.path.join(results_base, split_name)
            
            print(f"[{current_experiment}/{total_experiments}] {split_name}")
            
            # Run experiment
            result = run_single_experiment(model_name, train_split, experiment_dir)
            results_summary.append(result)
    
    # Print summary
    print("\n" + "="*70)
    print("PODSUMOWANIE SERII EKSPERYMENTÓW")
    print("="*70)
    print(f"\nCzas zakończenia: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Folder wyników: {results_base}\n")
    
    successful = sum(1 for r in results_summary if r['status'] == 'SUCCESS')
    failed = len(results_summary) - successful
    
    for i, result in enumerate(results_summary, 1):
        status_symbol = "✓" if result['status'] == 'SUCCESS' else "✗"
        print(f"{status_symbol} [{i}/{len(results_summary)}] {result['model']:15s} split_{result['split']} - {result['status']}")
        if 'error' in result:
            print(f"    Błąd: {result['error']}")
    
    print(f"\n{'='*70}")
    print(f"Razem eksperymentów: {len(results_summary)}")
    print(f"Sukcesów: {successful} ✓")
    print(f"Błędów: {failed} ✗")
    print(f"Wskaźnik sukcesu: {successful/len(results_summary)*100:.1f}%")
    print(f"{'='*70}\n")
    
    return results_summary


def main():
    """Main entry point"""
    print(f"\nGPU dostępne: {tf.config.list_physical_devices('GPU')}")
    
    # Configure experiments here
    experiments = {
        'model_names': ['simple_cnn','mobilenet', 'resnet'],  # Add 'mobilenet', 'resnet' ,simple_cnn
        'train_splits': [0.1, 0.3, 0.5, 0.7, 0.9],  
        'results_dir': 'results'
    }
    
    # Run all experiments
    results = run_experiments(**experiments)
    
    return results


if __name__ == "__main__":
    main()