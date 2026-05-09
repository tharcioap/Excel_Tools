# Random Forest Statistical Playground

This repository contains a complete VBA macro capable of transforming Excel into a Random Forest Statistical Playground using the powerful `Python in Excel` feature.

## Features
- **Zero-Installation Engine**: Relies on `Python in Excel` via the `=PY()` formula (powered by Anaconda).
- **Automated Configuration Panel**: Run the initial setup macro to generate a fully editable configuration tab (`RF_Config`).
- **Complete Flexibility**: Set Target Y, Features X, test splits, max depth, n_estimators, and select independent periods for comparison.
- **Robust Insights Output**: Automatically splits data, fills missing values, label-encodes categorical variables, fits models for separate periods, extracts Feature Importances and evaluates model performance (Accuracy for Classification, R2 for Regression).

## Usage
1. Open Excel and insert the VBA code located in `RandomForestPlayground.bas` into a new Module in the Visual Basic Editor (ALT + F11).
2. Run `SetupRandomForestPlayground`.
3. An `RF_Config` tab will be created. Point it to your raw dataset, set your features (X) and target (Y).
4. Click the `Run Random Forest` button generated in the tab.
5. The `RF_Insights` tab will display your statistical results and feature importances.
