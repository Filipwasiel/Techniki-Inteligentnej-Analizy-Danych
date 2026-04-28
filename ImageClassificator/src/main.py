# main.py
import config
from data_loader import load_data
from model import build_cnn_model
from evaluate import evaluate_and_plot

def main():
    print("1. Ładowanie danych z dysku...")
    train_ds, test_ds = load_data()

    print("2. Budowanie struktury modelu...")
    # Przekazujemy wymiary z pliku config
    input_shape = (config.IMG_SIZE[0], config.IMG_SIZE[1], 3)
    model = build_cnn_model(input_shape=input_shape)

    print("3. Rozpoczynanie trenowania...")
    history = model.fit(
        train_ds,
        epochs=config.EPOCHS,
        validation_data=test_ds
    )

    print("4. Generowanie raportów...")
    evaluate_and_plot(model, test_ds, history)
    print("Zakończono pomyślnie!")

if __name__ == "__main__":
    main()