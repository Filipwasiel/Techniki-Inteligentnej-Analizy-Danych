# Image Classificator

A binary image classification project using TensorFlow/Keras CNN to classify images (e.g., cats vs. dogs).

## 📋 Table of Contents

- [Overview](#overview)
- [Dataset](#dataset)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Model Architecture](#model-architecture)
- [Evaluation](#evaluation)
- [Testing](#testing)
- [Results](#results)

## 🎯 Overview

This project implements a Convolutional Neural Network (CNN) for binary image classification tasks. It includes:
- Data loading and preprocessing with TensorFlow
- CNN model training and evaluation
- Classification metrics and visualization
- Unit tests for model and data loader components

## 📊 Dataset

### Source

The dataset is obtained from **Kaggle**. You can download it from:
- **Kaggle Dataset URL**: https://www.kaggle.com/datasets/samuelcortinhas/cats-and-dogs-image-classification/data

### Directory Structure

After downloading, organize your data as follows:

```
data/
├── train/
│   ├── class_0/
│   │   ├── image_1.jpg
│   │   ├── image_2.jpg
│   │   └── ...
│   └── class_1/
│       ├── image_1.jpg
│       ├── image_2.jpg
│       └── ...
└── test/
    ├── class_0/
    │   ├── image_1.jpg
    │   └── ...
    └── class_1/
        ├── image_1.jpg
        └── ...
```

### Data Preparation

1. Download your dataset from Kaggle
2. Extract the files to the `data/` directory
3. Ensure images are organized in subdirectories by class
4. The data loader will automatically:
   - Resize images to 128x128 pixels
   - Normalize pixel values to [0, 1]
   - Create batches for efficient processing
   - Apply caching and prefetching for performance

## 📁 Project Structure

```
ImageClassificator/
├── src/
│   ├── __init__.py           # Package initialization
│   ├── main.py               # Main training script
│   ├── config.py             # Configuration parameters
│   ├── model.py              # CNN model definition
│   ├── data_loader.py        # Data loading and preprocessing
│   └── evaluate.py           # Evaluation and visualization
├── tests/
│   ├── model_test.py         # Model architecture tests
│   └── data_loader_test.py   # Data loader tests
├── data/                     # Dataset directory (not in repo)
│   ├── train/                # Training images
│   └── test/                 # Test images
└── README.md                 # This file
```

## 🚀 Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd ImageClassificator
   ```

2. **Create a virtual environment** (optional but recommended)
   ```bash
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
   
## 🏃 Usage

### Training the Model

Run the main training script:

```bash
cd src
python main.py
```


**Last Updated**: 2026-04-28
